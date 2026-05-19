    /**
     * 印刷系処理後も画面下部の印刷条件を保持する。
     *
     * 内容確認・印刷プレビュー・印刷を実行した後に、
     * 各種画面入力値が初期値に戻らないように Form / View に再設定する。
     *
     * @param request リクエスト情報
     * @param myForm  画面情報
     * @param myView  ビュー情報
     */
    private void keepPrintCondition(HttpServletRequest request,
            SampleForm myForm,
            Sample_View myView) {

        // 項目A
        String itemA = request.getParameter("item_a");
        if (itemA == null || "".equals(itemA)) {
            itemA = myForm.getItemA();
        }
        if (itemA == null || "".equals(itemA)) {
            itemA = myView.getItemA();
        }
        myForm.setItemA(itemA);
        myView.setItemA(itemA);

        // 項目B
        String itemB = request.getParameter("item_b");
        if (itemB == null || "".equals(itemB)) {
            itemB = myForm.getItemB();
        }
        if (itemB == null || "".equals(itemB)) {
            itemB = myView.getItemB();
        }
        myForm.setItemB(itemB);
        myView.setItemB(itemB);

        // 項目C
        String itemC = request.getParameter("item_c");
        if (itemC == null || "".equals(itemC)) {
            itemC = myForm.getItemC();
        }
        if (itemC == null || "".equals(itemC)) {
            itemC = myView.getItemC();
        }
        myForm.setItemC(itemC);
        myView.setItemC(itemC);

        // 項目D
        String itemD = request.getParameter("item_d");
        if (itemD == null || "".equals(itemD)) {
            itemD = myForm.getItemD();
        }
        if (itemD == null || "".equals(itemD)) {
            itemD = myView.getItemD();
        }
        myForm.setItemD(itemD);
        myView.setItemD(itemD);

        // 項目E
        String itemE = request.getParameter("item_e");
        if (itemE == null || "".equals(itemE)) {
            itemE = myForm.getItemE();
        }
        if (itemE == null || "".equals(itemE)) {
            itemE = myView.getItemE();
        }
        myForm.setItemE(itemE);
        myView.setItemE(itemE);

        // 項目F
        String itemF = request.getParameter("item_f");
        if (itemF == null || "".equals(itemF)) {
            itemF = myForm.getItemF();
        }
        if (itemF == null || "".equals(itemF)) {
            itemF = myView.getItemF();
        }
        myForm.setItemF(itemF);
        myView.setItemF(itemF);

        // 項目G
        String itemG = request.getParameter("item_g");
        if (itemG == null || "".equals(itemG)) {
            itemG = myForm.getItemG();
        }
        if (itemG == null || "".equals(itemG)) {
            itemG = myView.getItemG();
        }
        myForm.setItemG(itemG);
        myView.setItemG(itemG);
    }
