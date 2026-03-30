-- Migration to add formatting_time_seconds column to documents table

ALTER TABLE documents 
ADD COLUMN formatting_time FLOAT;

COMMENT ON COLUMN documents.formatting_time IS 'Duration of the formatting process in seconds';
