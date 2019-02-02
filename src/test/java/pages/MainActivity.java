package pages;

import helper.TestUtil;

public class MainActivity {

    public static void main(String as[]){
        Object[][] obj= TestUtil.getTestData("Sheet1" );
        System.out.println("obj.length = " + obj.length);

        for (Object[] sfsdf :obj)
        {
            for (Object ff:sfsdf) {
                System.out.print(ff+" || ");
            }
            System.out.println();
        }
    }
}
