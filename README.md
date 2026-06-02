-- 形式Aから形式Bへの変換処理
V_FORMAT_B_VAL := FUNC_A_TO_B(I_FIELD_VAL);

-- 逆変換によるバリデーション（整合性チェック）
IF V_FORMAT_B_VAL IS NOT NULL THEN
    V_REVERT_CHECK := FUNC_B_TO_A(V_FORMAT_B_VAL);
ELSE
    V_REVERT_CHECK := NULL;
END IF;

IF V_FORMAT_B_VAL IS NOT NULL
   AND V_REVERT_CHECK = I_FIELD_VAL THEN

    -- データが正しく変換できた場合の処理
    O_RESULT_B := V_FORMAT_B_VAL;
    O_RESULT_A := I_FIELD_VAL;

ELSE

    -- 変換エラーまたは不整合がある場合の処理
    O_RESULT_B := NULL;
    O_RESULT_A := I_RAW_DATA_ORG;

END IF;
