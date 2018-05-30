import sys, time
import win32com.client


class ConnectionManager:
    """ConnectionManager is a wrapper for the PCOMM.autECLConnMgr
    object."""

    def __init__(self):
        self.PCommConnMgr = win32com.client.Dispatch("PCOMM.autECLConnMgr")
        # self.OIA = win32com.client.Dispatch("PCOMM.autECLOIAObj")
        self.connList = self.PCommConnMgr.autECLConnList

        self.activeSession = None
        self.sessions = {}

    def get_available_connections(self):
        self.connList.Refresh()

        connections = []
        for i in range(self.connList.Count):
            if self.connList(i + 1).Ready:
                connections.append(self.connList(i + 1).Name)

        return connections

    def check_logged_in(self, session):
        if self.get_text(1, 36, 7, connection_name=session) == "Sign On":
            return False
        else:
            return True

    def open_session(self, connection_name):
        _session = win32com.client.Dispatch("PCOMM.autECLSession")
        _session.SetConnectionByName(connection_name)

        if not _session.Ready:
            _session.StartCommunication()

        self.sessions[connection_name] = _session

    def set_active_session(self, connection_name):
        self.activeSession = self.sessions[connection_name]

    def get_text(self, row, col, length=None, connection_name=None):
        result = None

        temp_session = self.activeSession
        if connection_name is not None:
            temp_session = self.sessions[connection_name]

        if length is None:
            temp_session.autECLPS.autECLFieldList.Refresh()
            field = temp_session.autECLPS.autECLFieldList.FindFieldByRowCol(row, col)
            length = field.Length

        result = temp_session.autECLPS.GetText(row, col, length)
        return result

    def send_keys(self, key, row=None, col=None, connection_name=None):
        temp_session = self.activeSession
        if connection_name is not None:
            temp_session = self.sessions[connection_name]

        if row is None or col is None:
            temp_session.autECLPS.SendKeys("%s" % key)

        else:
            temp_session.autECLPS.SendKeys("%s" % key, row, col)


    def enter(self):
        self.send_keys("[enter]")



# conn = ConnectionManager()
# d = conn.get_available_connections()
# for i in d:
#     conn.open_session(i)
#     conn.set_active_session(i)
#     conn.send_keys(1, "test " + str(i))