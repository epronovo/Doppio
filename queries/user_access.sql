-- ===========================================================================
-- USER ACCESS MATRIX  -  doppio.db
--
-- Lists every program in the new role model for one user, whether they have
-- access or not, and cross-checks each granting role against IFS.
-- Change the USID in the params line and run.
--
--   universe   : the 973 distinct FNIDs in ses400 (everything the load covers)
--   grants     : mns410 (user -> role) joined to ses400 (role -> program)
--   labels     : mns405.TX40 (description) and mns405.TX15 (old role code)
--   ifs        : ifs_user_roles matched to mns405.TX15 after normalising
--                (upper-case, underscores and spaces removed) - IFS writes the
--                old role as FGR300_RO where TX15 holds FGR300RO.
--                111 of the 153 roles have an IFS equivalent.
--   precedence : W beats R;  an N role alongside a grant is flagged as a
--                CONFLICT rather than silently resolved
--
-- access_level : WRITE / READ / DENIED / CONFLICT - deny + grant /
--                GRANTED - no options / NO ACCESS
-- ifs_check    : IFS OK             - user holds the matching IFS role
--                MISSING IN IFS     - role maps to IFS but user does not hold it
--                NO IFS EQUIVALENT  - no IFS role matches this M3 role
--                PARTIAL            - holds some but not all matching IFS roles
--                NOT AN IFS USER    - the USID has no IFS account (e.g. *ADMIN)
-- ===========================================================================

WITH params(usid) AS (VALUES ('AAMARTIN')),          -- <<< set the user here

-- every program in scope, with its M3 description
prog AS (
    SELECT DISTINCT s.FNID,
           (SELECT sec.Function_Description
              FROM security sec
             WHERE sec.FNID = s.FNID
             LIMIT 1)                                        AS description
      FROM ses400 s
),

-- new role  ->  matching IFS role name (NULL when IFS has no equivalent)
rolemap AS (
    SELECT m.ROLL, m.TX15, m.TX40, i.RoleName                AS ifs_role
      FROM mns405 m
      LEFT JOIN (SELECT DISTINCT RoleName FROM ifs_user_roles) i
             ON replace(replace(upper(trim(i.RoleName)), '_',''), ' ','')
              = replace(replace(upper(trim(m.TX15   )), '_',''), ' ','')
),

-- the IFS roles this particular user actually holds
ifs_held AS (
    SELECT DISTINCT RoleName
      FROM ifs_user_roles
     WHERE upper(UserAlias) = upper((SELECT usid FROM params))
),

-- does this USID exist in IFS at all?
is_ifs_user AS (
    SELECT COUNT(*) AS n
      FROM ifs_users
     WHERE upper(UserAlias) = upper((SELECT usid FROM params))
),

-- what this user is granted, rolled up to one row per program
granted AS (
    SELECT s.FNID,
           MAX(s.AL01)                                       AS can_create,
           MAX(s.AL02)                                       AS can_change,
           MAX(s.AL04)                                       AS can_delete,
           MAX(s.AL05)                                       AS can_display,
           MAX(CASE WHEN s.ROLL LIKE '%W' THEN 1 ELSE 0 END) AS has_w_role,
           MAX(CASE WHEN s.ROLL LIKE '%R' THEN 1 ELSE 0 END) AS has_r_role,
           MAX(CASE WHEN s.ROLL LIKE '%N' THEN 1 ELSE 0 END) AS has_n_role,
           COUNT(DISTINCT s.ROLL)                            AS role_count,
           GROUP_CONCAT(DISTINCT s.ROLL)                     AS roles,
           GROUP_CONCAT(DISTINCT rm.TX15)                    AS old_roles,
           GROUP_CONCAT(DISTINCT rm.TX40)                    AS role_descriptions,
           GROUP_CONCAT(DISTINCT rm.ifs_role)                AS ifs_roles,
           GROUP_CONCAT(DISTINCT h.RoleName)                 AS ifs_roles_held,
           COUNT(DISTINCT rm.ifs_role)                       AS ifs_mapped_cnt,
           COUNT(DISTINCT h.RoleName)                        AS ifs_held_cnt
      FROM mns410  u
      JOIN ses400  s  ON s.ROLL  = u.ROLL
      JOIN rolemap rm ON rm.ROLL = s.ROLL
      LEFT JOIN ifs_held h ON h.RoleName = rm.ifs_role
     WHERE u.USID = (SELECT usid FROM params)
     GROUP BY s.FNID
)

SELECT (SELECT usid FROM params)                             AS USID,
       p.FNID                                                AS program,
       p.description                                         AS program_description,
       CASE
         WHEN g.FNID IS NULL                        THEN 'NO ACCESS'
         WHEN g.has_n_role = 1
              AND (g.has_w_role = 1 OR g.has_r_role = 1)
                                                    THEN 'CONFLICT - deny + grant'
         WHEN g.has_n_role = 1                      THEN 'DENIED'
         WHEN g.can_create = 1
           OR g.can_change = 1
           OR g.can_delete = 1                      THEN 'WRITE'
         WHEN g.can_display = 1                     THEN 'READ'
         ELSE                                            'GRANTED - no options'
       END                                                   AS access_level,
       COALESCE(g.can_create , 0)                            AS opt1_create,
       COALESCE(g.can_change , 0)                            AS opt2_change,
       COALESCE(g.can_delete , 0)                            AS opt4_delete,
       COALESCE(g.can_display, 0)                            AS opt5_display,
       COALESCE(g.role_count , 0)                            AS roles_granting,
       g.roles                                               AS via_roles,
       g.old_roles                                           AS via_old_roles,     -- mns405.TX15
       g.role_descriptions                                   AS via_role_desc,     -- mns405.TX40
       g.ifs_roles                                           AS ifs_roles,         -- matching IFS role names
       g.ifs_roles_held                                      AS ifs_roles_held,    -- of those, held by this user
       CASE
         WHEN g.FNID IS NULL                          THEN NULL
         WHEN (SELECT n FROM is_ifs_user) = 0         THEN 'NOT AN IFS USER'
         WHEN g.ifs_mapped_cnt = 0                    THEN 'NO IFS EQUIVALENT'
         WHEN g.ifs_held_cnt = 0                      THEN 'MISSING IN IFS'
         WHEN g.ifs_held_cnt < g.ifs_mapped_cnt       THEN 'PARTIAL'
         ELSE                                              'IFS OK'
       END                                                   AS ifs_check
  FROM prog p
  LEFT JOIN granted g ON g.FNID = p.FNID
 ORDER BY CASE
            WHEN g.FNID IS NULL   THEN 4      -- no access last
            WHEN g.has_n_role = 1 THEN 1      -- denies and conflicts first
            WHEN g.can_create = 1 THEN 2
            ELSE 3
          END,
          p.FNID;


-- ===========================================================================
-- VARIANT A  -  only what the user CANNOT reach
--              add to the main query, before the ORDER BY:
--
--   WHERE g.FNID IS NULL OR g.has_n_role = 1
--
-- ===========================================================================


-- ===========================================================================
-- VARIANT B  -  role-level reconciliation:  M3 role  vs  IFS role, per user
--               one row per user + role, no program explosion
-- ===========================================================================

SELECT u.USID,
       u.ROLL                                                AS new_role,
       m.TX15                                                AS old_role,
       m.TX40                                                AS role_description,
       i.RoleName                                            AS ifs_role,
       CASE WHEN iu.UserAlias IS NULL THEN 'NOT AN IFS USER'
            WHEN i.RoleName  IS NULL  THEN 'NO IFS EQUIVALENT'
            WHEN h.RoleName  IS NULL  THEN 'MISSING IN IFS'
            ELSE                           'IFS OK'
       END                                                   AS ifs_check
  FROM mns410 u
  JOIN mns405 m ON m.ROLL = u.ROLL
  LEFT JOIN (SELECT DISTINCT RoleName FROM ifs_user_roles) i
         ON replace(replace(upper(trim(i.RoleName)),'_',''),' ','')
          = replace(replace(upper(trim(m.TX15   )),'_',''),' ','')
  LEFT JOIN ifs_users     iu ON upper(iu.UserAlias) = upper(u.USID)
  LEFT JOIN ifs_user_roles h ON upper(h.UserAlias)  = upper(u.USID)
                            AND h.RoleName          = i.RoleName
 WHERE u.USID = 'AAMARTIN'                                   -- <<< set the user
 ORDER BY ifs_check, u.ROLL;


-- ===========================================================================
-- VARIANT C  -  one-line summary per user, with the IFS reconciliation
-- ===========================================================================

WITH rolemap AS (
    SELECT m.ROLL, m.TX15, i.RoleName AS ifs_role
      FROM mns405 m
      LEFT JOIN (SELECT DISTINCT RoleName FROM ifs_user_roles) i
             ON replace(replace(upper(trim(i.RoleName)),'_',''),' ','')
              = replace(replace(upper(trim(m.TX15   )),'_',''),' ','')
)
SELECT u.USID,
       iu.EmailId,
       iu.Status                                                     AS ifs_status,
       iu.LastLoginDate,
       COUNT(DISTINCT u.ROLL)                                        AS m3_roles,
       COUNT(DISTINCT s.FNID)                                        AS programs_granted,
       973 - COUNT(DISTINCT s.FNID)                                  AS programs_no_access,
       COUNT(DISTINCT CASE WHEN s.AL01 = 1 THEN s.FNID END)          AS programs_write,
       COUNT(DISTINCT CASE WHEN rm.ifs_role IS NOT NULL
                            AND h.RoleName IS NULL
                           THEN u.ROLL END)                          AS roles_missing_in_ifs
  FROM mns410 u
  JOIN rolemap rm ON rm.ROLL = u.ROLL
  LEFT JOIN ses400        s  ON s.ROLL = u.ROLL
  LEFT JOIN ifs_users     iu ON upper(iu.UserAlias) = upper(u.USID)
  LEFT JOIN ifs_user_roles h ON upper(h.UserAlias)  = upper(u.USID)
                            AND h.RoleName          = rm.ifs_role
 GROUP BY u.USID, iu.EmailId, iu.Status, iu.LastLoginDate
 ORDER BY roles_missing_in_ifs DESC, programs_write DESC;
