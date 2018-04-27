package com.eastrobot.jacob;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class JacobUtil {
    // word文档
    private static Dispatch doc;
    // word运行程序对象
    private static ActiveXComponent wordApp;
    // 所有word文档集合
    private static Dispatch documents;
    // 选定的范围或插入点
    private static Dispatch selection;
    // 保存退出
    private static boolean saveOnExit = true;

    public static void main(String[] args) {
        // ComThread.InitSTA();
        String docStr = "C:\\Users\\User\\Desktop\\kbase-media-2003.doc";
        String anotherDocStr = "C:\\Users\\User\\Desktop\\test.doc";

        // MSWordManager wordManager = new MSWordManager(true);
        // wordManager.openDocument(docStr);
        // Dispatch newDoc = wordManager.createNewDocument();
        // wordManager.copyParagraphFromAnotherDoc(docStr, 1);

        wordApp = new ActiveXComponent("Word.Application");//word
        // // ActiveXComponent excelApp = new ActiveXComponent("Excel.Application");//excel
        // //设置后台静默处理。
        wordApp.setProperty("Visible", new Variant(false));
        //
        // // 获得文档集合
        documents = wordApp.getProperty("Documents").toDispatch();
        // // 打开文档
        Dispatch doc = Dispatch.call(documents, "Open", new Variant(docStr)).toDispatch();
        // // 当前页
        Dispatch browser = Dispatch.get(wordApp, "Browser").toDispatch();

        Dispatch selection = Dispatch.get(wordApp, "Selection").toDispatch();
        String line = Dispatch.call(selection, "information", 10).toString();
        String pages = Dispatch.call(selection, "Information", new Variant(4)).toString();
        System.out.println(pages);

        Dispatch.call(wordApp, "Close");
        Dispatch.call(wordApp, "Quit");
        // // // 下一页
        // // Dispatch.call(browser, "Next");
        //
        // Dispatch selection = Dispatch.get(wordApp, "Selection").toDispatch();
        // String pages = Dispatch.call(selection, "Information", new Variant(4)).toString();

        // ComThread.Release();
    }

}