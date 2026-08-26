-- 针对当前模式或指定模式批量编译
EXEC DBMS_UTILITY.COMPILE_SCHEMA(schema => 'YOUR_SCHEMA_NAME', compile_all => FALSE);
