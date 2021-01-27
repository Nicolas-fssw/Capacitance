from fbs_runtime.application_context.PyQt5 import ApplicationContext, cached_property
import sys
from Probe_E_Test_v01 import Cap_Test


class AppContext(ApplicationContext):
    def run(self):
        self.window.resize(1200, 800)
        self.window.show()
        return appctxt.app.exec_()

    def get_design(self):
        qtCreatorFile = self.get_resource("Probe_E_Tester_GUI_v01.ui")
        return qtCreatorFile

    def get_connections(self):
        connectionsFile = self.get_resource("connections_and_save.csv")
        return connectionsFile

    def get_config(self):
        configFile = self.get_resource("config.csv")
        return configFile

    @cached_property
    def window(self):
        return Cap_Test(self.get_design(), self.get_connections(), self.get_config())


if __name__ == "__main__":
    appctxt = AppContext()
    exit_code = appctxt.run()
    sys.exit(exit_code)
