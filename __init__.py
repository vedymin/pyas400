import sys, time
import win32com.client


class ConnectionManager:
    """ConnectionManager is a wrapper for the PCOMM.autECLConnMgr
    object."""

    def __init__(self):
        self.PCommConnMgr = win32com.client.Dispatch("PCOMM.autECLConnMgr")

        self.connList = self.PCommConnMgr.autECLConnList
        self.OIA = win32com.client.Dispatch("PCOMM.autECLOIA")

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
            if self.get_text(1, 2, 7, connection_name=session) == "MENUINI":
                self.send_keys("1")
                self.enter()
                self.enter()
                return True
            elif self.get_text(23, 17, 6, connection_name=session) == "HARDIS":
                self.enter()
                return True
            else:
                return True

    def open_session(self, connection_name):
        _session = win32com.client.Dispatch("PCOMM.autECLSession")
        _session.SetConnectionByName(connection_name)

        self.OIA.SetConnectionByName(connection_name)

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
            self.OIA.WaitForInputReady()
            temp_session.autECLPS.autECLFieldList.Refresh()
            field = temp_session.autECLPS.autECLFieldList.FindFieldByRowCol(row, col)
            length = field.Length

        self.OIA.WaitForInputReady()
        result = temp_session.autECLPS.GetText(row, col, length)
        return result

    def send_keys(self, key, row=None, col=None, connection_name=None):
        self.OIA.WaitForInputReady()

        temp_session = self.activeSession
        if connection_name is not None:
            temp_session = self.sessions[connection_name]

        if row is None or col is None:
            self.OIA.WaitForInputReady()
            temp_session.autECLPS.SendKeys("%s" % key)
            self.OIA.WaitForInputReady()

        else:
            self.OIA.WaitForInputReady()
            self.tab()
            self.OIA.WaitForInputReady()
            temp_session.autECLPS.SendKeys("%s" % key, row, col)
            self.OIA.WaitForInputReady()

    def enter(self):
        self.OIA.WaitForInputReady()
        self.send_keys("[enter]")
        self.OIA.WaitForInputReady()

    def set_cursor(self, row, col, connection_name=None):
        self.OIA.WaitForInputReady()

        temp_session = self.activeSession
        if connection_name is not None:
            temp_session = self.sessions[connection_name]

        self.OIA.WaitForInputReady()
        temp_session.autECLPS.SetCursorPos(row, col)
        self.OIA.WaitForInputReady()

    def tab(self, count=1, connection_name=None):
        self.OIA.WaitForInputReady()
        temp_session = self.activeSession
        if connection_name is not None:
            temp_session = self.sessions[connection_name]

        for n in range(count):
            self.OIA.WaitForInputReady()
            self.send_keys("[tab]")
            self.OIA.WaitForInputReady()

    def fkey(self, key, count=1, connection_name=None):
        self.OIA.WaitForInputReady()
        temp_session = self.activeSession
        if connection_name is not None:
            temp_session = self.sessions[connection_name]

        for n in range(count):
            self.OIA.WaitForInputReady()
            self.send_keys("[pf{}]".format(key))
            self.OIA.WaitForInputReady()

    def esc(self, count=1, connection_name=None):
        self.OIA.WaitForInputReady()
        temp_session = self.activeSession
        if connection_name is not None:
            temp_session = self.sessions[connection_name]

        for n in range(count):
            self.OIA.WaitForInputReady()
            self.send_keys("[attn]")
            time.sleep(1)
            self.OIA.WaitForInputReady()

# conn = ConnectionManager()
# d = conn.get_available_connections()
# for i in d:
#     conn.open_session(i)
#     conn.set_active_session(i)
#     conn.send_keys(1, "test " + str(i))
