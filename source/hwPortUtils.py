#hwPortUtils.py
#A part of NonVisual Desktop Access (NVDA)
#Copyright (C) 2001-2016 Chris Liechti, NV Access Limited
# Based on serial scanner code by Chris Liechti from https://raw.githubusercontent.com/pyserial/pyserial/81167536e796cc2e13aa16abd17a14634dc3aed1/pyserial/examples/scanwin32.py

"""Utilities for working with hardware connection ports.
"""

import itertools
import ctypes
from ctypes.wintypes import BOOL, WCHAR, HWND, DWORD, ULONG, WORD
import _winreg as winreg
from winKernel import SYSTEMTIME
import config
from logHandler import log

def ValidHandle(value):
	if value == 0:
		raise ctypes.WinError()
	return value

HDEVINFO = ctypes.c_void_p
PCWSTR = ctypes.c_wchar_p
HWND = ctypes.c_uint
PDWORD = ctypes.POINTER(DWORD)
ULONG_PTR = ctypes.POINTER(ULONG)
ULONGLONG = ctypes.c_ulonglong
NULL = 0

class GUID(ctypes.Structure):
	_fields_ = (
		('Data1', ctypes.c_ulong),
		('Data2', ctypes.c_ushort),
		('Data3', ctypes.c_ushort),
		('Data4', ctypes.c_ubyte*8),
	)
	def __str__(self):
		return "{%08x-%04x-%04x-%s-%s}" % (
			self.Data1,
			self.Data2,
			self.Data3,
			''.join(["%02x" % d for d in self.Data4[:2]]),
			''.join(["%02x" % d for d in self.Data4[2:]]),
		)

class SP_DEVINFO_DATA(ctypes.Structure):
	_fields_ = (
		('cbSize', DWORD),
		('ClassGuid', GUID),
		('DevInst', DWORD),
		('Reserved', ULONG_PTR),
	)
	def __str__(self):
		return "ClassGuid:%s DevInst:%s" % (self.ClassGuid, self.DevInst)
PSP_DEVINFO_DATA = ctypes.POINTER(SP_DEVINFO_DATA)

class SP_DEVICE_INTERFACE_DATA(ctypes.Structure):
	_fields_ = (
		('cbSize', DWORD),
		('InterfaceClassGuid', GUID),
		('Flags', DWORD),
		('Reserved', ULONG_PTR),
	)
	def __str__(self):
		return "InterfaceClassGuid:%s Flags:%s" % (self.InterfaceClassGuid, self.Flags)

PSP_DEVICE_INTERFACE_DATA = ctypes.POINTER(SP_DEVICE_INTERFACE_DATA)

PSP_DEVICE_INTERFACE_DETAIL_DATA = ctypes.c_void_p

class dummy(ctypes.Structure):
	_fields_=(("d1", DWORD), ("d2", WCHAR))
	_pack_ = 1
SIZEOF_SP_DEVICE_INTERFACE_DETAIL_DATA_W = ctypes.sizeof(dummy)

SetupDiDestroyDeviceInfoList = ctypes.windll.setupapi.SetupDiDestroyDeviceInfoList
SetupDiDestroyDeviceInfoList.argtypes = (HDEVINFO,)
SetupDiDestroyDeviceInfoList.restype = BOOL

SetupDiGetClassDevs = ctypes.windll.setupapi.SetupDiGetClassDevsW
SetupDiGetClassDevs.argtypes = (ctypes.POINTER(GUID), PCWSTR, HWND, DWORD)
SetupDiGetClassDevs.restype = ValidHandle # HDEVINFO

SetupDiEnumDeviceInterfaces = ctypes.windll.setupapi.SetupDiEnumDeviceInterfaces
SetupDiEnumDeviceInterfaces.argtypes = (HDEVINFO, PSP_DEVINFO_DATA, ctypes.POINTER(GUID), DWORD, PSP_DEVICE_INTERFACE_DATA)
SetupDiEnumDeviceInterfaces.restype = BOOL

SetupDiGetDeviceInterfaceDetail = ctypes.windll.setupapi.SetupDiGetDeviceInterfaceDetailW
SetupDiGetDeviceInterfaceDetail.argtypes = (HDEVINFO, PSP_DEVICE_INTERFACE_DATA, PSP_DEVICE_INTERFACE_DETAIL_DATA, DWORD, PDWORD, PSP_DEVINFO_DATA)
SetupDiGetDeviceInterfaceDetail.restype = BOOL

SetupDiGetDeviceRegistryProperty = ctypes.windll.setupapi.SetupDiGetDeviceRegistryPropertyW
SetupDiGetDeviceRegistryProperty.argtypes = (HDEVINFO, PSP_DEVINFO_DATA, DWORD, PDWORD, ctypes.c_void_p, DWORD, PDWORD)
SetupDiGetDeviceRegistryProperty.restype = BOOL

GUID_CLASS_COMPORT = GUID(0x86e0d1e0L, 0x8089, 0x11d0,
	(ctypes.c_ubyte*8)(0x9c, 0xe4, 0x08, 0x00, 0x3e, 0x30, 0x1f, 0x73))
GUID_DEVINTERFACE_USB_DEVICE = GUID(0xA5DCBF10, 0x6530, 0x11D2,
	(0x90, 0x1F, 0x00, 0xC0, 0x4F, 0xB9, 0x51, 0xED))

DIGCF_PRESENT = 2
DIGCF_DEVICEINTERFACE = 16
INVALID_HANDLE_VALUE = 0
ERROR_INSUFFICIENT_BUFFER = 122
SPDRP_HARDWAREID = 1
SPDRP_FRIENDLYNAME = 12
SPDRP_LOCATION_INFORMATION = 13
ERROR_NO_MORE_ITEMS = 259
DICS_FLAG_GLOBAL = 0x00000001
DIREG_DEV = 0x00000001

def _isDebug():
	return config.conf["debugLog"]["hwIo"]

def listComPorts(onlyAvailable=True):
	"""List com ports on the system.
	@param onlyAvailable: Only return ports that are currently available.
	@type onlyAvailable: bool
	@return: Generates dicts including keys of port, friendlyName and hardwareID.
	@rtype: generator of dict
	"""
	flags = DIGCF_DEVICEINTERFACE
	if onlyAvailable:
		flags |= DIGCF_PRESENT

	buf = ctypes.create_unicode_buffer(1024)
	g_hdi = SetupDiGetClassDevs(ctypes.byref(GUID_CLASS_COMPORT), None, NULL, flags)
	try:
		for dwIndex in xrange(256):
			entry = {}
			did = SP_DEVICE_INTERFACE_DATA()
			did.cbSize = ctypes.sizeof(did)

			if not SetupDiEnumDeviceInterfaces(
				g_hdi,
				None,
				ctypes.byref(GUID_CLASS_COMPORT),
				dwIndex,
				ctypes.byref(did)
			):
				if ctypes.GetLastError() != ERROR_NO_MORE_ITEMS:
					raise ctypes.WinError()
				break

			dwNeeded = DWORD()
			# get the size
			if not SetupDiGetDeviceInterfaceDetail(
				g_hdi,
				ctypes.byref(did),
				None, 0, ctypes.byref(dwNeeded),
				None
			):
				# Ignore ERROR_INSUFFICIENT_BUFFER
				if ctypes.GetLastError() != ERROR_INSUFFICIENT_BUFFER:
					raise ctypes.WinError()
			# allocate buffer
			class SP_DEVICE_INTERFACE_DETAIL_DATA_W(ctypes.Structure):
				_fields_ = (
					('cbSize', DWORD),
					('DevicePath', WCHAR*(dwNeeded.value - ctypes.sizeof(DWORD))),
				)
				def __str__(self):
					return "DevicePath:%s" % (self.DevicePath,)
			idd = SP_DEVICE_INTERFACE_DETAIL_DATA_W()
			idd.cbSize = SIZEOF_SP_DEVICE_INTERFACE_DETAIL_DATA_W
			devinfo = SP_DEVINFO_DATA()
			devinfo.cbSize = ctypes.sizeof(devinfo)
			if not SetupDiGetDeviceInterfaceDetail(
				g_hdi,
				ctypes.byref(did),
				ctypes.byref(idd), dwNeeded, None,
				ctypes.byref(devinfo)
			):
				raise ctypes.WinError()

			# hardware ID
			if not SetupDiGetDeviceRegistryProperty(
				g_hdi,
				ctypes.byref(devinfo),
				SPDRP_HARDWAREID,
				None,
				ctypes.byref(buf), ctypes.sizeof(buf) - 1,
				None
			):
				# Ignore ERROR_INSUFFICIENT_BUFFER
				if ctypes.GetLastError() != ERROR_INSUFFICIENT_BUFFER:
					raise ctypes.WinError()
			else:
				hwID = entry["hardwareID"] = buf.value

			regKey = ctypes.windll.setupapi.SetupDiOpenDevRegKey(g_hdi, ctypes.byref(devinfo), DICS_FLAG_GLOBAL, 0, DIREG_DEV, winreg.KEY_READ)
			try:
				try:
					port = entry["port"] = winreg.QueryValueEx(regKey, "PortName")[0]
				except WindowsError:
					# #6015: In some rare cases, this value doesn't exist.
					log.debugWarning("No PortName value for hardware ID %s" % hwID)
					continue
				if not port:
					log.debugWarning("Empty PortName value for hardware ID %s" % hwID)
					continue
				if hwID.startswith("BTHENUM\\"):
					# This is a Microsoft bluetooth port.
					try:
						addr = winreg.QueryValueEx(regKey, "Bluetooth_UniqueID")[0].split("#", 1)[1].split("_", 1)[0]
						addr = int(addr, 16)
						entry["bluetoothAddress"] = addr
						if addr:
							entry["bluetoothName"] = getBluetoothDeviceInfo(addr).szName
					except:
						pass
				elif hwID == r"Bluetooth\0004&0002":
					# This is a Toshiba bluetooth port.
					try:
						entry["bluetoothAddress"], entry["bluetoothName"] = getToshibaBluetoothPortInfo(port)
					except:
						pass
				elif hwID == r"{95C7A0A0-3094-11D7-A202-00508B9D7D5A}\BLUETOOTHPORT":
					try:
						entry["bluetoothAddress"], entry["bluetoothName"] = getWidcommBluetoothPortInfo(port)
					except:
						pass
			finally:
				ctypes.windll.advapi32.RegCloseKey(regKey)

			# friendly name
			if not SetupDiGetDeviceRegistryProperty(
				g_hdi,
				ctypes.byref(devinfo),
				SPDRP_FRIENDLYNAME,
				None,
				ctypes.byref(buf), ctypes.sizeof(buf) - 1,
				None
			):
				# #6007: SPDRP_FRIENDLYNAME sometimes doesn't exist/isn't valid.
				log.debugWarning("Couldn't get SPDRP_FRIENDLYNAME for %s: %s" % (port, ctypes.WinError()))
				entry["friendlyName"] = port
			else:
				entry["friendlyName"] = buf.value

			if _isDebug():
				log.debug("%r" % entry)
			yield entry

	finally:
		SetupDiDestroyDeviceInfoList(g_hdi)
	if _isDebug():
		log.debug("Finished listing com ports")

BLUETOOTH_MAX_NAME_SIZE = 248
BTH_ADDR = BLUETOOTH_ADDRESS = ULONGLONG

class BLUETOOTH_DEVICE_INFO(ctypes.Structure):
	_fields_ = (
		("dwSize", DWORD),
		("address", BLUETOOTH_ADDRESS),
		("ulClassofDevice", ULONG),
		("fConnected", BOOL),
		("fRemembered", BOOL),
		("fAuthenticated", BOOL),
		("stLastSeen", SYSTEMTIME),
		("stLastUsed", SYSTEMTIME),
		("szName", WCHAR * BLUETOOTH_MAX_NAME_SIZE)
	)
	def __init__(self, **kwargs):
		super(BLUETOOTH_DEVICE_INFO, self).__init__(dwSize=ctypes.sizeof(self), **kwargs)

def getBluetoothDeviceInfo(address):
	devInfo = BLUETOOTH_DEVICE_INFO(address=address)
	res = ctypes.windll["bthprops.cpl"].BluetoothGetDeviceInfo(None, ctypes.byref(devInfo))
	if res != 0:
		raise ctypes.WinError(res)
	return devInfo

def getToshibaBluetoothPortInfo(port):
	with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Toshiba\BluetoothStack\V1.0\EZC\DATA") as rootKey:
		for index in itertools.count():
			try:
				keyName = winreg.EnumKey(rootKey, index)
			except WindowsError:
				break
			with winreg.OpenKey(rootKey, keyName) as itemKey:
				with winreg.OpenKey(itemKey, "SCORIGINAL") as scorigKey:
					try:
						if winreg.QueryValueEx(scorigKey, "PORTNAME")[0].rstrip("\0") != port:
							# This isn't the port we're interested in.
							continue
					except WindowsError:
						# This isn't a COM port.
						continue
				addr = winreg.QueryValueEx(itemKey, "BDADDR")[0]
				# addr is a string of raw bytes.
				# Convert it to a single number.
				addr = sum(ord(byte) << (byteNum * 8) for byteNum, byte in enumerate(reversed(addr)))
				name = winreg.QueryValueEx(itemKey, "FRIENDLYNAME")[0].rstrip("\0")
				return addr, name
	raise LookupError

def getWidcommBluetoothPortInfo(port):
	with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Widcomm\BTConfig\AutoConnect") as rootKey:
		for index in itertools.count():
			try:
				keyName = winreg.EnumKey(rootKey, index)
			except WindowsError:
				break
			# The keys are the port number, but might be prefixed by 0s.
			# For example, COM4 is 0004.
			if keyName.lstrip("0") != port[3:]:
				# This isn't the port we're interested in.
				continue
			with winreg.OpenKey(rootKey, keyName) as itemKey:
				addr = winreg.QueryValueEx(itemKey, "BDAddress")[0]
				# addr is a string of raw bytes.
				# Convert it to a single number.
				addr = sum(ord(byte) << (byteNum * 8) for byteNum, byte in enumerate(reversed(addr)))
				name = winreg.QueryValueEx(itemKey, "BDName")[0]
				return addr, name
	raise LookupError

def listUsbDevices(onlyAvailable=True):
	"""List USB devices on the system.
	@param onlyAvailable: Only return devices that are currently available.
	@type onlyAvailable: bool
	@return: The USB vendor and product IDs in the form "VID_xxxx&PID_xxxx"
	@rtype: generator of unicode
	"""
	flags = DIGCF_DEVICEINTERFACE
	if onlyAvailable:
		flags |= DIGCF_PRESENT

	buf = ctypes.create_unicode_buffer(1024)
	g_hdi = SetupDiGetClassDevs(GUID_DEVINTERFACE_USB_DEVICE, None, NULL, flags)
	try:
		for dwIndex in xrange(256):
			did = SP_DEVICE_INTERFACE_DATA()
			did.cbSize = ctypes.sizeof(did)

			if not SetupDiEnumDeviceInterfaces(
				g_hdi,
				None,
				GUID_DEVINTERFACE_USB_DEVICE,
				dwIndex,
				ctypes.byref(did)
			):
				if ctypes.GetLastError() != ERROR_NO_MORE_ITEMS:
					raise ctypes.WinError()
				break

			dwNeeded = DWORD()
			# get the size
			if not SetupDiGetDeviceInterfaceDetail(
				g_hdi,
				ctypes.byref(did),
				None, 0, ctypes.byref(dwNeeded),
				None
			):
				# Ignore ERROR_INSUFFICIENT_BUFFER
				if ctypes.GetLastError() != ERROR_INSUFFICIENT_BUFFER:
					raise ctypes.WinError()
			# allocate buffer
			class SP_DEVICE_INTERFACE_DETAIL_DATA_W(ctypes.Structure):
				_fields_ = (
					('cbSize', DWORD),
					('DevicePath', WCHAR*(dwNeeded.value - ctypes.sizeof(DWORD))),
				)
				def __str__(self):
					return "DevicePath:%s" % (self.DevicePath,)
			idd = SP_DEVICE_INTERFACE_DETAIL_DATA_W()
			idd.cbSize = SIZEOF_SP_DEVICE_INTERFACE_DETAIL_DATA_W
			devinfo = SP_DEVINFO_DATA()
			devinfo.cbSize = ctypes.sizeof(devinfo)
			if not SetupDiGetDeviceInterfaceDetail(
				g_hdi,
				ctypes.byref(did),
				ctypes.byref(idd), dwNeeded, None,
				ctypes.byref(devinfo)
			):
				raise ctypes.WinError()

			# hardware ID
			if not SetupDiGetDeviceRegistryProperty(
				g_hdi,
				ctypes.byref(devinfo),
				SPDRP_HARDWAREID,
				None,
				ctypes.byref(buf), ctypes.sizeof(buf) - 1,
				None
			):
				# Ignore ERROR_INSUFFICIENT_BUFFER
				if ctypes.GetLastError() != ERROR_INSUFFICIENT_BUFFER:
					raise ctypes.WinError()
			else:
				# The string is of the form "usb\VID_xxxx&PID_xxxx&..."
				usbId = buf.value[4:21] # VID_xxxx&PID_xxxx
				if _isDebug():
					log.debug("%r" % usbId)
				yield usbId
	finally:
		SetupDiDestroyDeviceInfoList(g_hdi)
	if _isDebug():
		log.debug("Finished listing USB devices")

_hidGuid = None
def listHidDevices(onlyAvailable=True):
	"""List HID devices on the system.
	@param onlyAvailable: Only return devices that are currently available.
	@type onlyAvailable: bool
	@return: Generates dicts including keys such as hardwareID,
		usbID (in the form "VID_xxxx&PID_xxxx")
		and devicePath.
	@rtype: generator of dict
	"""
	global _hidGuid
	if not _hidGuid:
		_hidGuid = GUID()
		ctypes.windll.hid.HidD_GetHidGuid(ctypes.byref(_hidGuid))

	flags = DIGCF_DEVICEINTERFACE
	if onlyAvailable:
		flags |= DIGCF_PRESENT

	buf = ctypes.create_unicode_buffer(1024)
	g_hdi = SetupDiGetClassDevs(_hidGuid, None, NULL, flags)
	try:
		for dwIndex in xrange(256):
			did = SP_DEVICE_INTERFACE_DATA()
			did.cbSize = ctypes.sizeof(did)

			if not SetupDiEnumDeviceInterfaces(
				g_hdi,
				None,
				_hidGuid,
				dwIndex,
				ctypes.byref(did)
			):
				if ctypes.GetLastError() != ERROR_NO_MORE_ITEMS:
					raise ctypes.WinError()
				break

			dwNeeded = DWORD()
			# get the size
			if not SetupDiGetDeviceInterfaceDetail(
				g_hdi,
				ctypes.byref(did),
				None, 0, ctypes.byref(dwNeeded),
				None
			):
				# Ignore ERROR_INSUFFICIENT_BUFFER
				if ctypes.GetLastError() != ERROR_INSUFFICIENT_BUFFER:
					raise ctypes.WinError()
			# allocate buffer
			class SP_DEVICE_INTERFACE_DETAIL_DATA_W(ctypes.Structure):
				_fields_ = (
					('cbSize', DWORD),
					('DevicePath', WCHAR*(dwNeeded.value - ctypes.sizeof(DWORD))),
				)
				def __str__(self):
					return "DevicePath:%s" % (self.DevicePath,)
			idd = SP_DEVICE_INTERFACE_DETAIL_DATA_W()
			idd.cbSize = SIZEOF_SP_DEVICE_INTERFACE_DETAIL_DATA_W
			devinfo = SP_DEVINFO_DATA()
			devinfo.cbSize = ctypes.sizeof(devinfo)
			if not SetupDiGetDeviceInterfaceDetail(
				g_hdi,
				ctypes.byref(did),
				ctypes.byref(idd), dwNeeded, None,
				ctypes.byref(devinfo)
			):
				raise ctypes.WinError()

			# hardware ID
			if not SetupDiGetDeviceRegistryProperty(
				g_hdi,
				ctypes.byref(devinfo),
				SPDRP_HARDWAREID,
				None,
				ctypes.byref(buf), ctypes.sizeof(buf) - 1,
				None
			):
				# Ignore ERROR_INSUFFICIENT_BUFFER
				if ctypes.GetLastError() != ERROR_INSUFFICIENT_BUFFER:
					raise ctypes.WinError()
			else:
				hwId = buf.value
				info = {
					"hardwareID": hwId,
					"devicePath": idd.DevicePath}
				hwId = hwId.split("\\", 1)[1]
				if hwId.startswith("VID"):
					info["usbID"] = hwId[:17] # VID_xxxx&PID_xxxx
				if _isDebug():
					log.debug("%r" % info)
				yield info
	finally:
		SetupDiDestroyDeviceInfoList(g_hdi)
	if _isDebug():
		log.debug("Finished listing HID devices")
