-- ONLINE EXAM: TEMPORARY ADMIN LOGIN + SECURE EXCEL IMPORT
-- Run in Supabase SQL Editor as project owner.
-- Temporary username: MAHESH
-- Temporary password: mmm
-- Change the password hash before public launch.

CREATE TABLE IF NOT EXISTS public.exam_admin_users (
    username TEXT PRIMARY KEY,
    password_hash TEXT NOT NULL,
    is_active BOOLEAN NOT NULL DEFAULT TRUE,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS public.exam_admin_sessions (
    token_hash TEXT PRIMARY KEY,
    username TEXT NOT NULL REFERENCES public.exam_admin_users(username) ON DELETE CASCADE,
    expires_at TIMESTAMPTZ NOT NULL,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

ALTER TABLE public.exam_admin_users ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.exam_admin_sessions ENABLE ROW LEVEL SECURITY;
REVOKE ALL ON public.exam_admin_users, public.exam_admin_sessions FROM anon, authenticated, PUBLIC;

-- Seed or reset the temporary admin password to "mmm".
INSERT INTO public.exam_admin_users (username, password_hash, is_active)
VALUES ('MAHESH', crypt('mmm', gen_salt('bf')), TRUE)
ON CONFLICT (username)
DO UPDATE SET password_hash = EXCLUDED.password_hash, is_active = TRUE;

CREATE OR REPLACE FUNCTION public.exam_admin_login(
    p_username TEXT,
    p_password TEXT
)
RETURNS JSONB
LANGUAGE plpgsql
SECURITY DEFINER
SET search_path = ''
AS $$
DECLARE
    v_user public.exam_admin_users%ROWTYPE;
    v_token TEXT;
BEGIN
    IF NULLIF(BTRIM(p_username), '') IS NULL OR NULLIF(p_password, '') IS NULL THEN
        RETURN jsonb_build_object('success', FALSE, 'message', 'Enter username and password.');
    END IF;

    SELECT u.* INTO v_user
    FROM public.exam_admin_users AS u
    WHERE UPPER(u.username) = UPPER(BTRIM(p_username))
      AND u.is_active = TRUE;

    IF NOT FOUND OR crypt(p_password, v_user.password_hash) <> v_user.password_hash THEN
        -- Keep response generic to avoid revealing which field was wrong.
        RETURN jsonb_build_object('success', FALSE, 'message', 'Invalid username or password.');
    END IF;

    v_token := encode(gen_random_bytes(32), 'hex');

    INSERT INTO public.exam_admin_sessions (token_hash, username, expires_at)
    VALUES (
        encode(digest(v_token, 'sha256'), 'hex'),
        v_user.username,
        clock_timestamp() + INTERVAL '30 minutes'
    );

    RETURN jsonb_build_object(
        'success', TRUE,
        'session_token', v_token,
        'expires_in_seconds', 1800
    );
END;
$$;

CREATE OR REPLACE FUNCTION public.exam_admin_import_questions(
    p_session_token TEXT,
    p_exam_id TEXT,
    p_questions JSONB
)
RETURNS JSONB
LANGUAGE plpgsql
SECURITY DEFINER
SET search_path = ''
AS $$
DECLARE
    v_exam public.exam_details%ROWTYPE;
    v_count INTEGER;
    v_existing_count INTEGER;
    v_invalid_count INTEGER;
    v_inserted INTEGER;
    v_session_username TEXT;
    v_seen_ids INTEGER;
BEGIN
    IF NULLIF(p_session_token, '') IS NULL OR NULLIF(BTRIM(p_exam_id), '') IS NULL
       OR p_questions IS NULL OR jsonb_typeof(p_questions) <> 'array' THEN
        RETURN jsonb_build_object('success', FALSE, 'message', 'Invalid import request.');
    END IF;

    SELECT s.username INTO v_session_username
    FROM public.exam_admin_sessions AS s
    WHERE s.token_hash = encode(digest(p_session_token, 'sha256'), 'hex')
      AND s.expires_at > clock_timestamp();

    IF NOT FOUND THEN
        RETURN jsonb_build_object('success', FALSE, 'message', 'Admin session expired. Log in again.');
    END IF;

    SELECT e.* INTO v_exam
    FROM public.exam_details AS e
    WHERE e.exam_id = BTRIM(p_exam_id);

    IF NOT FOUND THEN
        RETURN jsonb_build_object('success', FALSE, 'message', 'Exam ID does not exist in exam_details.');
    END IF;

    v_count := jsonb_array_length(p_questions);
    IF v_count = 0 THEN
        RETURN jsonb_build_object('success', FALSE, 'message', 'No questions supplied.');
    END IF;

    IF v_count <> v_exam.max_marks THEN
        RETURN jsonb_build_object(
            'success', FALSE,
            'message', 'Question count must equal MAXMARKS.',
            'question_count', v_count,
            'max_marks', v_exam.max_marks
        );
    END IF;

    -- This initial importer only allows importing into an empty exam,
    -- avoiding accidental duplicate imports or partial replacements.
    SELECT COUNT(*)::INTEGER INTO v_existing_count
    FROM public.exam_questions AS q
    WHERE q.exam_id = v_exam.exam_id;

    IF v_existing_count > 0 THEN
        RETURN jsonb_build_object(
            'success', FALSE,
            'message', 'Questions already exist for this exam. Import cancelled; no rows changed.',
            'existing_count', v_existing_count
        );
    END IF;

    -- Validate required fields, matching exam ID, allowed sections and A-D keys.
    SELECT COUNT(*)::INTEGER INTO v_invalid_count
    FROM jsonb_array_elements(p_questions) AS x(item)
    WHERE NULLIF(BTRIM(x.item->>'question_id'), '') IS NULL
       OR NULLIF(BTRIM(x.item->>'section_id'), '') IS NULL
       OR NULLIF(BTRIM(x.item->>'qname'), '') IS NULL
       OR NULLIF(BTRIM(x.item->>'opt1'), '') IS NULL
       OR NULLIF(BTRIM(x.item->>'opt2'), '') IS NULL
       OR NULLIF(BTRIM(x.item->>'opt3'), '') IS NULL
       OR NULLIF(BTRIM(x.item->>'opt4'), '') IS NULL
       OR UPPER(BTRIM(x.item->>'answer_key')) NOT IN ('A', 'B', 'C', 'D')
       OR BTRIM(x.item->>'exam_id') <> v_exam.exam_id
       OR UPPER(BTRIM(x.item->>'section_id')) NOT IN ('MA', 'SCI', 'SS', 'MATH');

    IF v_invalid_count > 0 THEN
        RETURN jsonb_build_object(
            'success', FALSE,
            'message', 'One or more questions have invalid or missing fields.',
            'invalid_rows', v_invalid_count
        );
    END IF;

    -- Reject duplicate question IDs inside the upload.
    SELECT COUNT(*)::INTEGER - COUNT(DISTINCT BTRIM(x.item->>'question_id'))::INTEGER
    INTO v_seen_ids
    FROM jsonb_array_elements(p_questions) AS x(item);

    IF v_seen_ids > 0 THEN
        RETURN jsonb_build_object(
            'success', FALSE,
            'message', 'Duplicate QUESTIONID values found in the Excel file.'
        );
    END IF;

    -- Insert all rows in one transaction; any SQL error rolls back the entire call.
    INSERT INTO public.exam_questions (
        exam_id, question_id, section_id, qname,
        opt1, opt2, opt3, opt4, answer_key, hint
    )
    SELECT
        BTRIM(x.item->>'exam_id'),
        BTRIM(x.item->>'question_id'),
        UPPER(BTRIM(x.item->>'section_id')),
        BTRIM(x.item->>'qname'),
        BTRIM(x.item->>'opt1'),
        BTRIM(x.item->>'opt2'),
        BTRIM(x.item->>'opt3'),
        BTRIM(x.item->>'opt4'),
        UPPER(BTRIM(x.item->>'answer_key')),
        NULLIF(BTRIM(x.item->>'hint'), '')
    FROM jsonb_array_elements(p_questions) AS x(item);

    GET DIAGNOSTICS v_inserted = ROW_COUNT;

    -- Invalidate the used session after successful import.
    DELETE FROM public.exam_admin_sessions
    WHERE token_hash = encode(digest(p_session_token, 'sha256'), 'hex');

    RETURN jsonb_build_object(
        'success', TRUE,
        'message', 'Questions imported successfully.',
        'inserted_count', v_inserted,
        'exam_id', v_exam.exam_id
    );
END;
$$;

REVOKE ALL ON FUNCTION public.exam_admin_login(TEXT, TEXT) FROM PUBLIC;
REVOKE ALL ON FUNCTION public.exam_admin_import_questions(TEXT, TEXT, JSONB) FROM PUBLIC;

GRANT EXECUTE ON FUNCTION public.exam_admin_login(TEXT, TEXT) TO anon, authenticated;
GRANT EXECUTE ON FUNCTION public.exam_admin_import_questions(TEXT, TEXT, JSONB) TO anon, authenticated;
