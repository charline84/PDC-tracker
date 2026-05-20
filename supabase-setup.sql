-- Script de création de la table "permis_construire" pour Supabase
-- Copiez-collez ce code dans l'onglet "SQL Editor" de votre projet Supabase, puis cliquez sur "RUN".

CREATE TABLE permis_construire (
  id TEXT PRIMARY KEY,
  type TEXT,
  date_depot TEXT,
  date_obtention TEXT,
  statut TEXT,
  demandeur TEXT,
  siren_demandeur TEXT,
  adresse TEXT,
  description TEXT
);

-- Désactivation temporaire de la sécurité (RLS) pour permettre à GitHub Pages d'insérer des données
-- (Pour une application en production, il est recommandé de l'activer et de créer des "Policies")
ALTER TABLE permis_construire DISABLE ROW LEVEL SECURITY;
