package teprinciple.yang.list2excel.app;


import android.support.multidex.MultiDex;
import android.support.multidex.MultiDexApplication;


public class MyApp extends MultiDexApplication {
    @Override
    public void onCreate() {
        super.onCreate();
        MultiDex.install(this);
    }
}
