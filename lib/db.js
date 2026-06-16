require('dotenv').config();
const { createClient } = require('@supabase/supabase-js');

// service_role обходит RLS — используем для всех серверных операций
const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_SERVICE_KEY
);

// ── Projects ──────────────────────────────────────────────────────────────────

async function getProjects() {
  const { data, error } = await supabase
    .from('projects')
    .select('*')
    .order('created_at', { ascending: false });
  if (error) throw error;
  return data;
}

async function createProject({ name, description }) {
  const { data, error } = await supabase
    .from('projects')
    .insert({ name, description })
    .select()
    .single();
  if (error) throw error;
  return data;
}

async function deleteProject(id) {
  const { error } = await supabase.from('projects').delete().eq('id', id);
  if (error) throw error;
}

async function updateProject(id, { name, description }) {
  const fields = {};
  if (name        !== undefined) fields.name        = name;
  if (description !== undefined) fields.description = description;
  const { data, error } = await supabase
    .from('projects')
    .update(fields)
    .eq('id', id)
    .select()
    .single();
  if (error) throw error;
  return data;
}

// ── Analyses ──────────────────────────────────────────────────────────────────

async function getAnalyses(projectId) {
  const { data, error } = await supabase
    .from('analyses')
    .select(`*, files(*)`)
    .eq('project_id', projectId)
    .order('created_at', { ascending: false });
  if (error) throw error;
  return data;
}

async function createAnalysis({ projectId, filename }) {
  const { data, error } = await supabase
    .from('analyses')
    .insert({ project_id: projectId, filename, status: 'pending' })
    .select()
    .single();
  if (error) throw error;
  return data;
}

async function completeAnalysis(id, { statsTotal, statsAuto, statsRequest }) {
  const { data, error } = await supabase
    .from('analyses')
    .update({
      status: 'completed',
      stats_total:   statsTotal,
      stats_auto:    statsAuto,
      stats_request: statsRequest,
    })
    .eq('id', id)
    .select()
    .single();
  if (error) throw error;
  return data;
}

async function failAnalysis(id, errorMessage) {
  const { error } = await supabase
    .from('analyses')
    .update({ status: 'error', error_message: errorMessage })
    .eq('id', id);
  if (error) throw error;
}

// ── Files ─────────────────────────────────────────────────────────────────────

async function getFile(id) {
  const { data, error } = await supabase
    .from('files')
    .select('storage_path, file_type, analyses(filename)')
    .eq('id', id)
    .single();
  if (error) throw error;
  return data;
}

async function createFile({ analysisId, fileType, storagePath }) {
  const { data, error } = await supabase
    .from('files')
    .insert({ analysis_id: analysisId, file_type: fileType, storage_path: storagePath })
    .select()
    .single();
  if (error) throw error;
  return data;
}

// ── Storage ───────────────────────────────────────────────────────────────────

const BUCKET = 'analysis-files';

async function uploadFile(storagePath, buffer, contentType) {
  const { error } = await supabase.storage
    .from(BUCKET)
    .upload(storagePath, buffer, { contentType, upsert: true });
  if (error) throw error;
}

async function getSignedUrl(storagePath, expiresIn = 3600) {
  const { data, error } = await supabase.storage
    .from(BUCKET)
    .createSignedUrl(storagePath, expiresIn);
  if (error) throw error;
  return data.signedUrl;
}

async function downloadFile(storagePath) {
  const { data, error } = await supabase.storage
    .from(BUCKET)
    .download(storagePath);
  if (error) throw error;
  return Buffer.from(await data.arrayBuffer());
}

module.exports = {
  getProjects, createProject, updateProject, deleteProject,
  getAnalyses, createAnalysis, completeAnalysis, failAnalysis,
  getFile, createFile,
  uploadFile, getSignedUrl, downloadFile,
};
