CREATE TABLE IF NOT EXISTS students (
    id BIGSERIAL PRIMARY KEY,
    full_name TEXT NOT NULL,
    student_group TEXT NOT NULL,
    username TEXT NOT NULL UNIQUE,
    email TEXT UNIQUE,
    password_hash TEXT NOT NULL,
    role TEXT NOT NULL DEFAULT 'student' CHECK (role IN ('student', 'admin')),
    created_at TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP
);

ALTER TABLE students ADD COLUMN IF NOT EXISTS username TEXT;
ALTER TABLE students ADD COLUMN IF NOT EXISTS role TEXT NOT NULL DEFAULT 'student';
UPDATE students SET username = 'student_' || id WHERE username IS NULL;
ALTER TABLE students ALTER COLUMN username SET NOT NULL;
ALTER TABLE students ALTER COLUMN email DROP NOT NULL;
CREATE UNIQUE INDEX IF NOT EXISTS students_username_key ON students (username);

CREATE TABLE IF NOT EXISTS lessons (
    id BIGSERIAL UNIQUE NOT NULL,
    slug TEXT PRIMARY KEY,
    title TEXT NOT NULL,
    body TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS questions (
    id BIGSERIAL PRIMARY KEY,
    lesson_slug TEXT NOT NULL REFERENCES lessons(slug) ON DELETE CASCADE,
    prompt TEXT NOT NULL,
    options JSONB NOT NULL,
    correct_index INTEGER NOT NULL CHECK (correct_index >= 0),
    question_type TEXT NOT NULL DEFAULT 'single',
    correct_answers JSONB NOT NULL DEFAULT '[]'::jsonb,
    UNIQUE (lesson_slug, prompt)
);

ALTER TABLE questions ADD COLUMN IF NOT EXISTS question_type TEXT NOT NULL DEFAULT 'single';
ALTER TABLE questions ADD COLUMN IF NOT EXISTS correct_answers JSONB NOT NULL DEFAULT '[]'::jsonb;
UPDATE questions
SET correct_answers = jsonb_build_array(correct_index)
WHERE question_type = 'single' AND correct_answers = '[]'::jsonb;

CREATE TABLE IF NOT EXISTS test_attempts (
    id BIGSERIAL PRIMARY KEY,
    student_id BIGINT NOT NULL REFERENCES students(id) ON DELETE CASCADE,
    lesson_slug TEXT NOT NULL REFERENCES lessons(slug),
    correct_count INTEGER NOT NULL,
    question_count INTEGER NOT NULL,
    percentage INTEGER NOT NULL,
    grade SMALLINT NOT NULL CHECK (grade BETWEEN 2 AND 5),
    created_at TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS practice_tasks (
    id BIGSERIAL UNIQUE NOT NULL,
    slug TEXT PRIMARY KEY,
    title TEXT NOT NULL,
    instructions TEXT NOT NULL,
    extensions TEXT[] NOT NULL,
    rubric JSONB
);

CREATE TABLE IF NOT EXISTS practice_submissions (
    id BIGSERIAL PRIMARY KEY,
    student_id BIGINT NOT NULL REFERENCES students(id) ON DELETE CASCADE,
    task_slug TEXT NOT NULL REFERENCES practice_tasks(slug),
    relative_path TEXT NOT NULL,
    status TEXT NOT NULL,
    grade SMALLINT CHECK (grade BETWEEN 2 AND 5),
    created_at TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP
);

ALTER TABLE practice_submissions ADD COLUMN IF NOT EXISTS variant_number INTEGER
    CHECK (variant_number BETWEEN 1 AND 30);
ALTER TABLE practice_submissions ADD COLUMN IF NOT EXISTS file_name TEXT;
ALTER TABLE practice_submissions ADD COLUMN IF NOT EXISTS file_data BYTEA;
ALTER TABLE practice_submissions ALTER COLUMN relative_path DROP NOT NULL;

CREATE TABLE IF NOT EXISTS practice_variant_assignments (
    id BIGSERIAL PRIMARY KEY,
    student_id BIGINT NOT NULL REFERENCES students(id) ON DELETE CASCADE,
    task_slug TEXT NOT NULL REFERENCES practice_tasks(slug),
    variant_number INTEGER NOT NULL CHECK (variant_number BETWEEN 1 AND 30),
    assigned_at TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP,
    UNIQUE (student_id, task_slug),
    UNIQUE (task_slug, variant_number)
);

ALTER TABLE practice_variant_assignments
    DROP CONSTRAINT IF EXISTS practice_variant_assignments_task_slug_variant_number_key;

UPDATE practice_submissions
SET variant_number = NULL
WHERE task_slug = 'practice-2' AND variant_number IS NOT NULL;

DELETE FROM practice_variant_assignments
WHERE task_slug = 'practice-2';
