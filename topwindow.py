import win32gui
import re, sys
import win32con
import win32api
import time


class WindowMgr():
	'''Encapsulates some calls to the winapi for windows management'''
	def __init__(self):
		'''Constructor'''
		self._handle = None
		
	def find_window(self, window_name, class_name = None):
		'''find a window by its window_name'''
		self._handle = win32gui.FindWindow(class_name, window_name)
		
	def __window_enum_callback(self, hwnd, wildcard):
		'''pass to win32gui.EnumWindows() to check all the opened windows'''
		# if re.match(wildcard, str(win32gui.GetWindowText(hwnd))) != None:
			# self._handle = hwnd    
		win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
		win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0,0,0,0, win32con.SWP_NOMOVE | win32con.SWP_NOACTIVATE| win32con.SWP_NOOWNERZORDER|win32con.SWP_SHOWWINDOW)
	
	def find_window_wildcard(self, wildcard):
		self._handle = None
		win32gui.EnumWindows(self.__window_enum_callback, wildcard)		
	
	def set_foreground(self):
		'''put the window in the forground'''
		win32gui.SetForegroundWindow(self._handle)
		
# def main():
	# w = WindowMgr()
	# w.find_window_wildcard('.*Users\\aaa*.')
	# w.set_foreground()
	
if __name__ == '__main__':
	hWndList = []
	win32gui.EnumWindows(lambda hWnd, param: param.append(hWnd), hWndList)
	for hwnd in hWndList:
		text = win32gui.GetWindowText(hwnd)
		if text.find("Notepad++") >= 0:
			print(text)
			break
	win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0,0,1500,800, win32con.SWP_NOMOVE | win32con.SWP_NOACTIVATE| win32con.SWP_NOOWNERZORDER|win32con.SWP_SHOWWINDOW)
	
	time.sleep(5)

