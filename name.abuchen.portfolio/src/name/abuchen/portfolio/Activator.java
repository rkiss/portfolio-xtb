package name.abuchen.portfolio;

import org.apache.logging.log4j.LogManager;
import org.osgi.framework.BundleActivator;
import org.osgi.framework.BundleContext;

public class Activator implements BundleActivator
{

    @Override
    public void start(BundleContext context) throws Exception
    {
        // 🚨 CRITICAL: force Log4j initialization in single thread
        LogManager.getContext(false);
    }

    @Override
    public void stop(BundleContext context) throws Exception
    {
    }

}
