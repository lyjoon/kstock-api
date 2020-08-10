package com.github.kstock.api.kh;

import com.github.kstock.TestApp;
import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import org.junit.Test;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * ocx 연동을 위한 jacob TC
 */
public class TestJacob extends TestApp {
    public static void Main(String[] args) {
        Logger log = LoggerFactory.getLogger(TestJacob.class);
        log.info("testJacob2");
        System.out.println("testJacob2");
    }

    @Test
    public void testJacob(){
        log.info("testJacob");
        //ActiveXComponent xl = new ActiveXComponent("{A1574A0D-6BFA-4BD7-9020-DED88711818D}");
        ActiveXComponent xl = new ActiveXComponent("KHOPENAPI.KHOpenAPICtrl");
        //ActiveXComponent xl = new ActiveXComponent(new Dispatch("KHOPENAPI.KHOpenAPICtrl"));
        //ActiveXComponent xl = new ActiveXComponent("Excel.Application");

        Object xlo = xl.getObject();
        try {
            log.info("version : {}", xl.getProperty("Version"));
            xl.setProperty("Visible", new Variant(true));
            Dispatch workbooks = xl.getProperty("Workbooks").toDispatch();
            Dispatch workbook = Dispatch.get(workbooks,"Add").toDispatch();
            Dispatch sheet = Dispatch.get(workbook,"ActiveSheet").toDispatch();
            Dispatch a1 = Dispatch.invoke(sheet, "Range", Dispatch.Get,
                    new Object[] {"A1"},
                    new int[1]).toDispatch();
            Dispatch a2 = Dispatch.invoke(sheet, "Range", Dispatch.Get,
                    new Object[] {"A2"},
                    new int[1]).toDispatch();
            Dispatch.put(a1, "Value", "123.456");
            Dispatch.put(a2, "Formula", "=A1*2");
            log.info("a1 from excel : {}", Dispatch.get(a1, "Value"));
            log.info("a2 from excel : {}", Dispatch.get(a2, "Value"));
            Variant f = new Variant(false);
            Dispatch.call(workbook, "Close", f);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            xl.invoke("Quit", new Variant[] {});
        }
    }
}
