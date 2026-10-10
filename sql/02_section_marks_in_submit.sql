-- Return authoritative section-wise correct counts in the submission response.
-- Run as project owner. Do not submit real student attempts until authorization is added.

CREATE OR REPLACE FUNCTION public.exam_submit_and_score(p_attempt_id uuid)
 RETURNS jsonb
 LANGUAGE plpgsql
 SECURITY DEFINER
 SET search_path TO ''
AS $function$
DECLARE
    v_attempt public.exam_attempts%ROWTYPE;
    v_exam public.exam_details%ROWTYPE;
    v_result_id UUID;
    v_question_count INTEGER;
    v_correct INTEGER;
    v_wrong INTEGER;
    v_unanswered INTEGER;
    v_total INTEGER;
    v_submitted_at TIMESTAMPTZ := clock_timestamp();
    v_final_status TEXT;
BEGIN
    -- 1. Lock the attempt so it cannot be finalized twice.
    SELECT a.*
    INTO v_attempt
    FROM public.exam_attempts AS a
    WHERE a.attempt_id = p_attempt_id
    FOR UPDATE;

    IF NOT FOUND THEN
        RETURN JSONB_BUILD_OBJECT(
            'success', FALSE,
            'message', 'Invalid exam attempt.'
        );
    END IF;

    -- 2. Return the existing result if already finalized.
    IF v_attempt.status IN ('submitted', 'expired') THEN
        SELECT r.result_id INTO v_result_id
        FROM public.exam_results AS r
        WHERE r.attempt_id = p_attempt_id;

        IF FOUND THEN
            RETURN JSONB_BUILD_OBJECT(
                'success', TRUE,
                'message', 'Exam was already finalized.',
                'result_id', v_result_id,
                'exam_id', v_attempt.exam_id,
                'total_marks', (SELECT r.total_marks FROM public.exam_results r WHERE r.result_id = v_result_id),
                'max_marks', (SELECT r.max_marks FROM public.exam_results r WHERE r.result_id = v_result_id),
                'correct_count', (SELECT r.correct_count FROM public.exam_results r WHERE r.result_id = v_result_id),
                'wrong_count', (SELECT r.wrong_count FROM public.exam_results r WHERE r.result_id = v_result_id),
                'unanswered_count', (SELECT r.unanswered_count FROM public.exam_results r WHERE r.result_id = v_result_id),
                'section_marks', (
                    SELECT COALESCE(jsonb_agg(jsonb_build_object(
                        'section_id', rs.section_id,
                        'correct_count', rs.correct_count
                    ) ORDER BY rs.section_id), '[]'::jsonb)
                    FROM public.exam_result_sections rs
                    WHERE rs.result_id = v_result_id
                ),
                'status', v_attempt.status
            );
        END IF;
        RETURN JSONB_BUILD_OBJECT('success', FALSE, 'message', 'Finalized result record missing.');
    END IF;

    -- 3. Load exam configuration.
    SELECT e.*
    INTO v_exam
    FROM public.exam_details AS e
    WHERE e.exam_id = v_attempt.exam_id;

    IF NOT FOUND THEN
        RETURN JSONB_BUILD_OBJECT(
            'success', FALSE,
            'message', 'Exam configuration not found.'
        );
    END IF;

    -- 4. Check that questions exist and match max marks.
    SELECT COUNT(*)::INTEGER
    INTO v_question_count
    FROM public.exam_questions AS q
    WHERE q.exam_id = v_attempt.exam_id;

    IF v_question_count = 0 THEN
        RETURN JSONB_BUILD_OBJECT(
            'success', FALSE,
            'message', 'No questions are available for this exam.'
        );
    END IF;

    IF v_question_count <> v_exam.max_marks THEN
        RETURN JSONB_BUILD_OBJECT(
            'success', FALSE,
            'message', 'Question count must equal the configured maximum marks.',
            'question_count', v_question_count,
            'configured_max_marks', v_exam.max_marks
        );
    END IF;

    IF EXISTS (
        SELECT 1 FROM public.exam_questions q
        WHERE q.exam_id = v_attempt.exam_id
          AND UPPER(BTRIM(q.answer_key)) NOT IN ('A', 'B', 'C', 'D')
    ) THEN
        RETURN JSONB_BUILD_OBJECT('success', FALSE,
            'message', 'Invalid answer key found. Expected A, B, C, or D.');
    END IF;

    -- 5. Score saved answers.
    -- Supports answer keys OPT1-OPT4, OPTION1-OPTION4,
    -- A-D, 1-4, or the exact option text.
    WITH scored AS (
        SELECT
            q.id,
            q.section_id,
            sa.selected_option,
            CASE
                WHEN UPPER(REGEXP_REPLACE(
                    BTRIM(q.answer_key), '\\s+', '', 'g'
                )) IN ('OPT1', 'OPTION1', 'A', '1')
                    THEN 'OPT1'
                WHEN UPPER(REGEXP_REPLACE(
                    BTRIM(q.answer_key), '\\s+', '', 'g'
                )) IN ('OPT2', 'OPTION2', 'B', '2')
                    THEN 'OPT2'
                WHEN UPPER(REGEXP_REPLACE(
                    BTRIM(q.answer_key), '\\s+', '', 'g'
                )) IN ('OPT3', 'OPTION3', 'C', '3')
                    THEN 'OPT3'
                WHEN UPPER(REGEXP_REPLACE(
                    BTRIM(q.answer_key), '\\s+', '', 'g'
                )) IN ('OPT4', 'OPTION4', 'D', '4')
                    THEN 'OPT4'
                WHEN LOWER(BTRIM(q.answer_key)) = LOWER(BTRIM(q.opt1))
                    THEN 'OPT1'
                WHEN LOWER(BTRIM(q.answer_key)) = LOWER(BTRIM(q.opt2))
                    THEN 'OPT2'
                WHEN LOWER(BTRIM(q.answer_key)) = LOWER(BTRIM(q.opt3))
                    THEN 'OPT3'
                WHEN LOWER(BTRIM(q.answer_key)) = LOWER(BTRIM(q.opt4))
                    THEN 'OPT4'
                ELSE NULL
            END AS correct_option
        FROM public.exam_questions AS q
        LEFT JOIN public.student_answers AS sa
          ON sa.question_row_id = q.id
         AND sa.attempt_id = p_attempt_id
        WHERE q.exam_id = v_attempt.exam_id
    )
    SELECT
        COUNT(*) FILTER (
            WHERE selected_option IS NOT NULL
              AND selected_option = correct_option
        )::INTEGER,
        COUNT(*) FILTER (
            WHERE selected_option IS NOT NULL
              AND selected_option <> correct_option
        )::INTEGER,
        COUNT(*) FILTER (
            WHERE selected_option IS NULL
        )::INTEGER
    INTO v_correct, v_wrong, v_unanswered
    FROM scored;

    v_total := v_correct;

    -- 6. Determine whether the attempt ended by timeout.
    IF v_attempt.status = 'expired'
       OR v_submitted_at >= v_attempt.deadline_at THEN
        v_final_status := 'expired';
    ELSE
        v_final_status := 'submitted';
    END IF;

    -- 7. Store the final overall result.
    INSERT INTO public.exam_results (
        attempt_id,
        exam_id,
        student_id,
        max_marks,
        total_marks,
        wrong_count,
        correct_count,
        unanswered_count,
        submitted_at
    )
    VALUES (
        p_attempt_id,
        v_attempt.exam_id,
        v_attempt.student_id,
        v_exam.max_marks,
        v_total,
        v_wrong,
        v_correct,
        v_unanswered,
        v_submitted_at
    )
    RETURNING result_id INTO v_result_id;

    -- 8. Store marks for every section dynamically.
    WITH scored AS (
        SELECT
            q.section_id,
            sa.selected_option,
            CASE
                WHEN UPPER(REGEXP_REPLACE(
                    BTRIM(q.answer_key), '\\s+', '', 'g'
                )) IN ('OPT1', 'OPTION1', 'A', '1')
                    THEN 'OPT1'
                WHEN UPPER(REGEXP_REPLACE(
                    BTRIM(q.answer_key), '\\s+', '', 'g'
                )) IN ('OPT2', 'OPTION2', 'B', '2')
                    THEN 'OPT2'
                WHEN UPPER(REGEXP_REPLACE(
                    BTRIM(q.answer_key), '\\s+', '', 'g'
                )) IN ('OPT3', 'OPTION3', 'C', '3')
                    THEN 'OPT3'
                WHEN UPPER(REGEXP_REPLACE(
                    BTRIM(q.answer_key), '\\s+', '', 'g'
                )) IN ('OPT4', 'OPTION4', 'D', '4')
                    THEN 'OPT4'
                WHEN LOWER(BTRIM(q.answer_key)) = LOWER(BTRIM(q.opt1))
                    THEN 'OPT1'
                WHEN LOWER(BTRIM(q.answer_key)) = LOWER(BTRIM(q.opt2))
                    THEN 'OPT2'
                WHEN LOWER(BTRIM(q.answer_key)) = LOWER(BTRIM(q.opt3))
                    THEN 'OPT3'
                WHEN LOWER(BTRIM(q.answer_key)) = LOWER(BTRIM(q.opt4))
                    THEN 'OPT4'
                ELSE NULL
            END AS correct_option
        FROM public.exam_questions AS q
        LEFT JOIN public.student_answers AS sa
          ON sa.question_row_id = q.id
         AND sa.attempt_id = p_attempt_id
        WHERE q.exam_id = v_attempt.exam_id
    )
    INSERT INTO public.exam_result_sections (
        result_id,
        section_id,
        max_marks,
        total_marks,
        wrong_count,
        correct_count,
        unanswered_count
    )
    SELECT
        v_result_id,
        section_id,
        COUNT(*)::INTEGER,
        COUNT(*) FILTER (
            WHERE selected_option IS NOT NULL
              AND selected_option = correct_option
        )::INTEGER,
        COUNT(*) FILTER (
            WHERE selected_option IS NOT NULL
              AND selected_option <> correct_option
        )::INTEGER,
        COUNT(*) FILTER (
            WHERE selected_option IS NOT NULL
              AND selected_option = correct_option
        )::INTEGER,
        COUNT(*) FILTER (
            WHERE selected_option IS NULL
        )::INTEGER
    FROM scored
    GROUP BY section_id;

    -- 9. Finalize the attempt.
    UPDATE public.exam_attempts
    SET status = v_final_status,
        submitted_at = v_submitted_at
    WHERE attempt_id = p_attempt_id;

    -- 10. Return a summary.
    RETURN JSONB_BUILD_OBJECT(
        'success', TRUE,
        'message', 'Exam scored and finalized successfully.',
        'result_id', v_result_id,
        'exam_id', v_attempt.exam_id,
        'max_marks', v_exam.max_marks,
        'correct_count', v_correct,
        'wrong_count', v_wrong,
        'unanswered_count', v_unanswered,
        'total_marks', v_total,
        'section_marks', (
            SELECT COALESCE(jsonb_agg(jsonb_build_object(
                'section_id', rs.section_id,
                'correct_count', rs.correct_count
            ) ORDER BY rs.section_id), '[]'::jsonb)
            FROM public.exam_result_sections rs
            WHERE rs.result_id = v_result_id
        ),
        'status', v_final_status,
        'submitted_at', v_submitted_at
    );
END;
$function$;
