-- Run this in Supabase SQL Editor FIRST to wipe everything clean
-- Go to: Supabase Dashboard → SQL Editor → paste this → Run

DROP TABLE IF EXISTS sessions CASCADE;
DROP TABLE IF EXISTS user_projects CASCADE;
DROP TABLE IF EXISTS columns_config CASCADE;
DROP TABLE IF EXISTS records CASCADE;
DROP TABLE IF EXISTS dropdown_lists CASCADE;
DROP TABLE IF EXISTS logos CASCADE;
DROP TABLE IF EXISTS doc_types CASCADE;
DROP TABLE IF EXISTS users CASCADE;
DROP TABLE IF EXISTS projects CASCADE;
DROP TABLE IF EXISTS settings CASCADE;
DROP TABLE IF EXISTS _v5_project CASCADE;
DROP TABLE IF EXISTS _v5_doc_types CASCADE;
DROP TABLE IF EXISTS _v5_records CASCADE;
DROP TABLE IF EXISTS _v5_columns_config CASCADE;
DROP TABLE IF EXISTS _v5_dropdown_lists CASCADE;
DROP TABLE IF EXISTS _v5_logos CASCADE;
DROP TABLE IF EXISTS _v5_settings CASCADE;
DROP TABLE IF EXISTS project CASCADE;

-- Confirm
SELECT 'All tables dropped successfully ✅' as result;
