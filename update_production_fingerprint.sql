-- SQL script to update HR Manager fingerprint in production database
-- Run this in Render.com PostgreSQL console

-- First, check current fingerprint
SELECT 
    u.username,
    u.employee_name,
    u.department,
    ed.browser_fingerprint,
    LENGTH(ed.browser_fingerprint) as fp_length
FROM "user" u
JOIN employee_data ed ON u.id = ed.user_id
WHERE u.username = 'hr_user' OR u.department = 'HR';

-- Update fingerprint to the reference value
UPDATE employee_data 
SET browser_fingerprint = '396520d70ea1f79dd21caffd85085795'
WHERE user_id = (
    SELECT id FROM "user" 
    WHERE username = 'hr_user' OR department = 'HR' 
    LIMIT 1
);

-- Verify the update
SELECT 
    u.username,
    u.employee_name,
    ed.browser_fingerprint,
    CASE 
        WHEN ed.browser_fingerprint = '396520d70ea1f79dd21caffd85085795' THEN 'MATCH ✅'
        ELSE 'MISMATCH ❌'
    END as status
FROM "user" u
JOIN employee_data ed ON u.id = ed.user_id
WHERE u.username = 'hr_user' OR u.department = 'HR';

