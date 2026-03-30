-- Migration: add tracked_changes_url column to documents table

ALTER TABLE documents
ADD COLUMN IF NOT EXISTS tracked_changes_url TEXT;

COMMENT ON COLUMN documents.tracked_changes_url IS 'Storage path to the tracked-changes .docx file for this document';
