-- =====================================================================
-- GPS SOURCE CENSUS  --  disaster 4882DR
--
-- Purpose: answer "where are the damage coordinates actually coming
-- from?" before building any extraction. Counts every place a damage
-- coordinate can live, so we can see which sources carry real volume
-- and which are noise.
--
-- Reads only. Safe to run repeatedly.
--
-- Source tiers:
--   T1 structured, user-entered in a GPS field (per damage category)
--   T2 structured but ambiguous -- dbo_applicant_damage.Latitude/Longitude
--      sits beside Address_1/City/Zip + Geocode_Confidence_Id +
--      UnGeocodeable, so it is probably a GEOCODE of the damage address,
--      not a hand-keyed GPS. Section 3 tests that.
--   T3 free text, needs regex
--
-- SCHEMA TRAP: the category answer tables disagree on the name of the
-- first coordinate. catc/catf/cata_removal use Start_Latitude; catd/
-- cate/catg use plain Latitude for the same concept. Both then use
-- End_Latitude for the second point.
--
-- Western-hemisphere assumption: the T3 regex requires a negative
-- longitude. Fine for Region V, revisit if this ever runs outside CONUS.
-- =====================================================================

-- Damages in scope for this disaster. Note: joined via Applicant_Id
-- directly (not through applicant_project) and NO county join -- the
-- org->county table is many-to-many and fans rows out up to 11x.
WITH dmg AS (
  SELECT
      ad.Applicant_Damage_Id,
      ad.Applicant_Project_Id,
      ad.Work_Category_Id,
      ad.Latitude,
      ad.Longitude,
      ad.Geocode_Confidence_Id,
      ad.UnGeocodeable,
      ad.Name              AS Damage_Name,
      ad.Damage_Description,
      ad.Address_1, ad.City, ad.Zip
  FROM fac_trax.odp.dbo_applicant_damage ad
  JOIN fac_trax.odp.dbo_applicant a ON ad.Applicant_Id = a.Applicant_Id
  JOIN fac_trax.odp.dbo_event     e ON e.Event_Id      = a.Event_Id
  WHERE e.Job_Number = '4882DR'
),

-- ---------------------------------------------------------------
-- SECTION 1 -- how many damages exist at all, by work category
-- ---------------------------------------------------------------
totals AS (
  SELECT
      '0. TOTAL damages in disaster' AS source,
      'n/a'                          AS geom_role,
      COUNT(*)                       AS coord_rows,
      COUNT(*)                       AS damages
  FROM dmg
),

-- ---------------------------------------------------------------
-- SECTION 2 -- T1: structured coordinates on the category answer
-- tables. These are the ones the current GPSExtraction query never
-- touches, because it INNER JOINs only _generaldamage.
-- ---------------------------------------------------------------
structured AS (
  SELECT 'catA_removal.Start' AS source, 'line_start' AS geom_role,
         COUNT(*) AS coord_rows, COUNT(DISTINCT d.Applicant_Damage_Id) AS damages
    FROM dmg d JOIN fac_trax.odp.dbo_applicant_damage_cata_removal t
      ON t.Applicant_Damage_Id = d.Applicant_Damage_Id
   WHERE t.Start_Latitude IS NOT NULL AND t.Start_Longitude IS NOT NULL
  UNION ALL
  SELECT 'catA_removal.End', 'line_end', COUNT(*), COUNT(DISTINCT d.Applicant_Damage_Id)
    FROM dmg d JOIN fac_trax.odp.dbo_applicant_damage_cata_removal t
      ON t.Applicant_Damage_Id = d.Applicant_Damage_Id
   WHERE t.End_Latitude IS NOT NULL AND t.End_Longitude IS NOT NULL
  UNION ALL
  SELECT 'catA_disposal', 'point', COUNT(*), COUNT(DISTINCT d.Applicant_Damage_Id)
    FROM dmg d JOIN fac_trax.odp.dbo_applicant_damage_cata_disposal t
      ON t.Applicant_Damage_Id = d.Applicant_Damage_Id
   WHERE t.Latitude IS NOT NULL AND t.Longitude IS NOT NULL
  UNION ALL
  SELECT 'catA_tempsr', 'point', COUNT(*), COUNT(DISTINCT d.Applicant_Damage_Id)
    FROM dmg d JOIN fac_trax.odp.dbo_applicant_damage_cata_tempsr t
      ON t.Applicant_Damage_Id = d.Applicant_Damage_Id
   WHERE t.Latitude IS NOT NULL AND t.Longitude IS NOT NULL
  UNION ALL
  -- CAT C = ROADS AND BRIDGES. This is the one that matters most.
  SELECT 'catC_answer.Start', 'line_start', COUNT(*), COUNT(DISTINCT d.Applicant_Damage_Id)
    FROM dmg d JOIN fac_trax.odp.dbo_applicant_damage_catc_answer t
      ON t.Applicant_Damage_Id = d.Applicant_Damage_Id
   WHERE t.Start_Latitude IS NOT NULL AND t.Start_Longitude IS NOT NULL
  UNION ALL
  SELECT 'catC_answer.End', 'line_end', COUNT(*), COUNT(DISTINCT d.Applicant_Damage_Id)
    FROM dmg d JOIN fac_trax.odp.dbo_applicant_damage_catc_answer t
      ON t.Applicant_Damage_Id = d.Applicant_Damage_Id
   WHERE t.End_Latitude IS NOT NULL AND t.End_Longitude IS NOT NULL
  UNION ALL
  SELECT 'catD_answer.Start', 'line_start', COUNT(*), COUNT(DISTINCT d.Applicant_Damage_Id)
    FROM dmg d JOIN fac_trax.odp.dbo_applicant_damage_catd_answer t
      ON t.Applicant_Damage_Id = d.Applicant_Damage_Id
   WHERE t.Latitude IS NOT NULL AND t.Longitude IS NOT NULL
  UNION ALL
  SELECT 'catD_answer.End', 'line_end', COUNT(*), COUNT(DISTINCT d.Applicant_Damage_Id)
    FROM dmg d JOIN fac_trax.odp.dbo_applicant_damage_catd_answer t
      ON t.Applicant_Damage_Id = d.Applicant_Damage_Id
   WHERE t.End_Latitude IS NOT NULL AND t.End_Longitude IS NOT NULL
  UNION ALL
  SELECT 'catE_answer.Start', 'line_start', COUNT(*), COUNT(DISTINCT d.Applicant_Damage_Id)
    FROM dmg d JOIN fac_trax.odp.dbo_applicant_damage_cate_answer t
      ON t.Applicant_Damage_Id = d.Applicant_Damage_Id
   WHERE t.Latitude IS NOT NULL AND t.Longitude IS NOT NULL
  UNION ALL
  SELECT 'catE_answer.End', 'line_end', COUNT(*), COUNT(DISTINCT d.Applicant_Damage_Id)
    FROM dmg d JOIN fac_trax.odp.dbo_applicant_damage_cate_answer t
      ON t.Applicant_Damage_Id = d.Applicant_Damage_Id
   WHERE t.End_Latitude IS NOT NULL AND t.End_Longitude IS NOT NULL
  UNION ALL
  SELECT 'catF_answer.Start', 'line_start', COUNT(*), COUNT(DISTINCT d.Applicant_Damage_Id)
    FROM dmg d JOIN fac_trax.odp.dbo_applicant_damage_catf_answer t
      ON t.Applicant_Damage_Id = d.Applicant_Damage_Id
   WHERE t.Start_Latitude IS NOT NULL AND t.Start_Longitude IS NOT NULL
  UNION ALL
  SELECT 'catF_answer.End', 'line_end', COUNT(*), COUNT(DISTINCT d.Applicant_Damage_Id)
    FROM dmg d JOIN fac_trax.odp.dbo_applicant_damage_catf_answer t
      ON t.Applicant_Damage_Id = d.Applicant_Damage_Id
   WHERE t.End_Latitude IS NOT NULL AND t.End_Longitude IS NOT NULL
  UNION ALL
  SELECT 'catG_answer.Start', 'line_start', COUNT(*), COUNT(DISTINCT d.Applicant_Damage_Id)
    FROM dmg d JOIN fac_trax.odp.dbo_applicant_damage_catg_answer t
      ON t.Applicant_Damage_Id = d.Applicant_Damage_Id
   WHERE t.Latitude IS NOT NULL AND t.Longitude IS NOT NULL
  UNION ALL
  SELECT 'catG_answer.End', 'line_end', COUNT(*), COUNT(DISTINCT d.Applicant_Damage_Id)
    FROM dmg d JOIN fac_trax.odp.dbo_applicant_damage_catg_answer t
      ON t.Applicant_Damage_Id = d.Applicant_Damage_Id
   WHERE t.End_Latitude IS NOT NULL AND t.End_Longitude IS NOT NULL
),

-- ---------------------------------------------------------------
-- SECTION 3 -- T2: the ambiguous one
-- ---------------------------------------------------------------
ambiguous AS (
  SELECT 'applicant_damage.Lat/Lon (T2 ambiguous)' AS source, 'point' AS geom_role,
         COUNT(*) AS coord_rows, COUNT(DISTINCT Applicant_Damage_Id) AS damages
    FROM dmg
   WHERE Latitude IS NOT NULL AND Longitude IS NOT NULL
),

-- ---------------------------------------------------------------
-- SECTION 4 -- T3: free text. Loose pair pattern: a lat-ish number,
-- any junk separator (comma/tab/space/slash/"GPS:"), then a negative
-- lon-ish number. Tolerates "- 86.805" (space after the sign).
-- ---------------------------------------------------------------
freetext AS (
  SELECT 'TEXT generaldamage.Location_Grouping' AS source, 'text' AS geom_role,
         COUNT(*) AS coord_rows, COUNT(DISTINCT d.Applicant_Damage_Id) AS damages
    FROM dmg d JOIN fac_trax.odp.dbo_applicant_damage_generaldamage g
      ON g.Applicant_Damage_Id = d.Applicant_Damage_Id
   WHERE g.Location_Grouping RLIKE '[0-9]{1,2}\\.[0-9]{3,}\\s*[,;/|]?\\s*-\\s*[0-9]{2,3}\\.[0-9]{3,}'
  UNION ALL
  SELECT 'TEXT generaldamage.Location_Grouping_Root', 'text', COUNT(*), COUNT(DISTINCT d.Applicant_Damage_Id)
    FROM dmg d JOIN fac_trax.odp.dbo_applicant_damage_generaldamage g
      ON g.Applicant_Damage_Id = d.Applicant_Damage_Id
   WHERE g.Location_Grouping_Root RLIKE '[0-9]{1,2}\\.[0-9]{3,}\\s*[,;/|]?\\s*-\\s*[0-9]{2,3}\\.[0-9]{3,}'
  UNION ALL
  SELECT 'TEXT generaldamage.Damage_Description', 'text', COUNT(*), COUNT(DISTINCT d.Applicant_Damage_Id)
    FROM dmg d JOIN fac_trax.odp.dbo_applicant_damage_generaldamage g
      ON g.Applicant_Damage_Id = d.Applicant_Damage_Id
   WHERE g.Damage_Description RLIKE '[0-9]{1,2}\\.[0-9]{3,}\\s*[,;/|]?\\s*-\\s*[0-9]{2,3}\\.[0-9]{3,}'
  UNION ALL
  SELECT 'TEXT generaldamage.Component_Description', 'text', COUNT(*), COUNT(DISTINCT d.Applicant_Damage_Id)
    FROM dmg d JOIN fac_trax.odp.dbo_applicant_damage_generaldamage g
      ON g.Applicant_Damage_Id = d.Applicant_Damage_Id
   WHERE g.Component_Description RLIKE '[0-9]{1,2}\\.[0-9]{3,}\\s*[,;/|]?\\s*-\\s*[0-9]{2,3}\\.[0-9]{3,}'
  UNION ALL
  SELECT 'TEXT applicant_damage.Damage_Description', 'text', COUNT(*), COUNT(DISTINCT Applicant_Damage_Id)
    FROM dmg
   WHERE Damage_Description RLIKE '[0-9]{1,2}\\.[0-9]{3,}\\s*[,;/|]?\\s*-\\s*[0-9]{2,3}\\.[0-9]{3,}'
  UNION ALL
  SELECT 'TEXT applicant_damage.Name', 'text', COUNT(*), COUNT(DISTINCT Applicant_Damage_Id)
    FROM dmg
   WHERE Damage_Name RLIKE '[0-9]{1,2}\\.[0-9]{3,}\\s*[,;/|]?\\s*-\\s*[0-9]{2,3}\\.[0-9]{3,}'
  UNION ALL
  SELECT 'TEXT catC_answer.Location_Description', 'text', COUNT(*), COUNT(DISTINCT d.Applicant_Damage_Id)
    FROM dmg d JOIN fac_trax.odp.dbo_applicant_damage_catc_answer t
      ON t.Applicant_Damage_Id = d.Applicant_Damage_Id
   WHERE t.Location_Description RLIKE '[0-9]{1,2}\\.[0-9]{3,}\\s*[,;/|]?\\s*-\\s*[0-9]{2,3}\\.[0-9]{3,}'
  UNION ALL
  SELECT 'TEXT catC_answer.Facility_Description', 'text', COUNT(*), COUNT(DISTINCT d.Applicant_Damage_Id)
    FROM dmg d JOIN fac_trax.odp.dbo_applicant_damage_catc_answer t
      ON t.Applicant_Damage_Id = d.Applicant_Damage_Id
   WHERE t.Facility_Description RLIKE '[0-9]{1,2}\\.[0-9]{3,}\\s*[,;/|]?\\s*-\\s*[0-9]{2,3}\\.[0-9]{3,}'
  UNION ALL
  SELECT 'TEXT catC_answer.Repair_Comments', 'text', COUNT(*), COUNT(DISTINCT d.Applicant_Damage_Id)
    FROM dmg d JOIN fac_trax.odp.dbo_applicant_damage_catc_answer t
      ON t.Applicant_Damage_Id = d.Applicant_Damage_Id
   WHERE t.Repair_Comments RLIKE '[0-9]{1,2}\\.[0-9]{3,}\\s*[,;/|]?\\s*-\\s*[0-9]{2,3}\\.[0-9]{3,}'
  UNION ALL
  SELECT 'TEXT catA_removal.Location', 'text', COUNT(*), COUNT(DISTINCT d.Applicant_Damage_Id)
    FROM dmg d JOIN fac_trax.odp.dbo_applicant_damage_cata_removal t
      ON t.Applicant_Damage_Id = d.Applicant_Damage_Id
   WHERE t.Location RLIKE '[0-9]{1,2}\\.[0-9]{3,}\\s*[,;/|]?\\s*-\\s*[0-9]{2,3}\\.[0-9]{3,}'
  UNION ALL
  SELECT 'TEXT catB_epm.Location', 'text', COUNT(*), COUNT(DISTINCT d.Applicant_Damage_Id)
    FROM dmg d JOIN fac_trax.odp.dbo_applicant_damage_catb_epm t
      ON t.Applicant_Damage_Id = d.Applicant_Damage_Id
   WHERE t.Location RLIKE '[0-9]{1,2}\\.[0-9]{3,}\\s*[,;/|]?\\s*-\\s*[0-9]{2,3}\\.[0-9]{3,}'
  UNION ALL
  SELECT 'TEXT applicant_damage_note.Note', 'text', COUNT(*), COUNT(DISTINCT d.Applicant_Damage_Id)
    FROM dmg d JOIN fac_trax.odp.dbo_applicant_damage_note t
      ON t.Applicant_Damage_Id = d.Applicant_Damage_Id
   WHERE t.Note RLIKE '[0-9]{1,2}\\.[0-9]{3,}\\s*[,;/|]?\\s*-\\s*[0-9]{2,3}\\.[0-9]{3,}'
)

SELECT * FROM totals
UNION ALL SELECT * FROM structured
UNION ALL SELECT * FROM ambiguous
UNION ALL SELECT * FROM freetext
ORDER BY damages DESC, source;


-- =====================================================================
-- FOLLOW-UP PROBES -- run these separately, each answers one question
-- =====================================================================

-- Q1. What does GPS_Type_Lookup_Id mean? No dbo_gps_type* lookup table
-- exists in odp, so the domain has to be inferred. If it separates
-- "single point" from "line segment", it decides whether a Cat C road
-- damage gets checked at one point or sampled along its length.
-- SELECT GPS_Type_Lookup_Id, COUNT(*) AS n,
--        SUM(CASE WHEN End_Latitude IS NOT NULL THEN 1 ELSE 0 END) AS has_end
--   FROM fac_trax.odp.dbo_applicant_damage_catc_answer
--  GROUP BY GPS_Type_Lookup_Id ORDER BY n DESC;

-- Q2. Is applicant_damage.Lat/Lon a geocode or a hand-keyed GPS?
-- SELECT * FROM fac_trax.odp.dbo_geocode_confidence;
-- Then: how many damages share an identical coordinate with another
-- damage from the same applicant? A repeated exact pair across many
-- damages = one org-level address geocoded over and over, not a site.
-- SELECT Applicant_Id, Latitude, Longitude, COUNT(*) AS damages_at_this_exact_point
--   FROM fac_trax.odp.dbo_applicant_damage
--  WHERE Latitude IS NOT NULL GROUP BY 1,2,3 HAVING COUNT(*) > 1
--  ORDER BY 4 DESC LIMIT 50;

-- Q3. Is `ssz` a copy of `odp`? information_schema shows 7409 vs 7414
-- columns -- near-identical. Need to know which is authoritative before
-- any of this gets scheduled.
-- SELECT COUNT(*) FROM fac_trax.ssz.dbo_applicant_damage;
