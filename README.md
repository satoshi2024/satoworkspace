SELECT 
    s.sid,
    s.serial#,
    s.username,
    s.status,
    s.machine,
    s.program,
    s.sql_id,
    q.sql_text,
    'ALTER SYSTEM KILL SESSION ''' || s.sid || ',' || s.serial# || ''' IMMEDIATE;' AS kill_sql
FROM 
    v$session s
LEFT JOIN 
    v$sql q ON s.sql_id = q.sql_id
WHERE 
    s.type != 'BACKGROUND'
    AND s.status = 'ACTIVE' -- 如果要查所有连接（含空闲），可注释掉此行
ORDER BY 
    s.logon_time DESC;
