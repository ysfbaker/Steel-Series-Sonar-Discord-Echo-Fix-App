"""
SteelSeries Sonar Discord Echo Fix
Development by CodeXart Studio's
"""

import threading, sys, os, winreg, ctypes, ctypes.wintypes, uuid, psutil, unicodedata
import pystray
from PIL import Image, ImageDraw

APP_NAME    = "SteelSeries Sonar Discord Echo Fix"
APP_VERSION = "1.0.0"

SONAR_KEYWORDS = [
	"kulakl",
	"head",
	"steelseries sonar - microphone",
	"sonar microphone",
]

TARGET_DEVICE_KEYWORDS = [
	"arctis nova 7 gen 2",
	"steelseries sonar - microphone",
]

ole32 = ctypes.windll.ole32
ole32.CoInitialize.argtypes  = [ctypes.c_void_p]
ole32.CoInitialize.restype   = ctypes.HRESULT
ole32.CoUninitialize.restype = None
ole32.CoCreateInstance.argtypes = [ctypes.c_void_p, ctypes.c_void_p, ctypes.c_ulong,  ctypes.c_void_p, ctypes.c_void_p]
ole32.CoCreateInstance.restype = ctypes.HRESULT

CLSCTX_ALL = 23

def _guid(s):
	b = uuid.UUID(s).bytes_le
	return (ctypes.c_byte * 16)(*b)

CLSID_MMDeviceEnumerator = _guid("{BCDE0395-E52F-467C-8E3D-C4579291692E}")
IID_IMMDeviceEnumerator  = _guid("{A95664D2-9614-4F35-A746-DE8DB63617E6}")
IID_IMMDeviceCollection  = _guid("{0BD7A1BE-7A1A-44DB-8397-CC5392387B5E}")
IID_IMMDevice            = _guid("{D666063F-1587-4E43-81F1-B948E807363F}")
IID_IAudioSessionManager2   = _guid("{77AA99A0-1BD6-484F-8BC7-2C654C9A9B6F}")
IID_IAudioSessionEnumerator = _guid("{E2F5BB11-0570-40CA-ACDD-3AA01277DEE8}")
IID_IAudioSessionControl2   = _guid("{BFB7FF88-7239-4FC9-8FA2-07C950BE9C6D}")
IID_ISimpleAudioVolume      = _guid("{87CE5498-68D6-44E5-9215-6DA47EF883D8}")
IID_IPropertyStore          = _guid("{886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99}")

# PKEY_Device_FriendlyName  {A45C254E-DF1C-4EFD-8020-67D146A850E0} pid=14
PKEY_FN_GUID = _guid("{A45C254E-DF1C-4EFD-8020-67D146A850E0}")
PKEY_FN_PID  = 14

# VTable index constants (0-based, after QueryInterface/AddRef/Release = 0,1,2)
# IMMDeviceEnumerator
VTBL_Enumerator_EnumAudioEndpoints = 3   # index 3
VTBL_Enumerator_GetDefaultAudioEndpoint = 4

# IMMDeviceCollection
VTBL_Collection_GetCount = 3
VTBL_Collection_Item     = 4

# IMMDevice
VTBL_Device_Activate          = 3
VTBL_Device_OpenPropertyStore = 4
VTBL_Device_GetId             = 5

# IPropertyStore
VTBL_PropStore_GetCount = 3
VTBL_PropStore_GetAt    = 4
VTBL_PropStore_GetValue = 5

# IAudioSessionManager2
VTBL_ASM2_GetSessionEnumerator = 5

# IAudioSessionEnumerator
VTBL_ASE_GetCount   = 3
VTBL_ASE_GetSession = 4

# IAudioSessionControl2
VTBL_ASC2_GetProcessId = 14   # 3(IUnk)+9(ASC)+2 = 14

# ISimpleAudioVolume
VTBL_SAV_SetMasterVolume = 3
VTBL_SAV_GetMasterVolume = 4
VTBL_SAV_SetMute         = 5
VTBL_SAV_GetMute         = 6


def _vtbl_call(com_ptr, index, restype, *argtypes_and_args):
	"""Call a COM vtable method by index."""
	# argtypes_and_args: list of (ctype, value) pairs
	argtypes = [ctypes.c_void_p] + [a[0] for a in argtypes_and_args]
	args     = [com_ptr]         + [a[1] for a in argtypes_and_args]
	vt = ctypes.cast(com_ptr, ctypes.POINTER(ctypes.c_void_p))
	vt = ctypes.cast(vt[0],   ctypes.POINTER(ctypes.c_void_p))
	fn = ctypes.cast(vt[index], ctypes.CFUNCTYPE(restype, *argtypes))
	return fn(*args)


def _qi(com_ptr, iid):
	"""QueryInterface — returns new void* or None."""
	out = ctypes.c_void_p(0)
	hr  = _vtbl_call(com_ptr, 0, ctypes.HRESULT, ((ctypes.c_byte*16), iid), (ctypes.POINTER(ctypes.c_void_p), ctypes.byref(out)))
	return out if (hr == 0 and out) else None


def _release(com_ptr):
	if com_ptr:
		_vtbl_call(com_ptr, 2, ctypes.c_ulong)


def _create_device_enumerator():
	obj = ctypes.c_void_p(0)
	hr  = ole32.CoCreateInstance(
		CLSID_MMDeviceEnumerator, None, CLSCTX_ALL,
		IID_IMMDeviceEnumerator, ctypes.byref(obj))
	return obj if hr == 0 else None


def _enum_audio_endpoints(enumerator, data_flow: int = 2):
	"""Returns IMMDeviceCollection* or None. data_flow: 0=render,1=capture,2=all; stateMask=1(active)"""
	col = ctypes.c_void_p(0)
	# HRESULT EnumAudioEndpoints(EDataFlow, DWORD, IMMDeviceCollection**)
	vt  = ctypes.cast(enumerator, ctypes.POINTER(ctypes.c_void_p))
	vt  = ctypes.cast(vt[0],      ctypes.POINTER(ctypes.c_void_p))
	fn  = ctypes.cast(vt[VTBL_Enumerator_EnumAudioEndpoints], ctypes.CFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p, ctypes.c_uint, ctypes.c_uint, ctypes.POINTER(ctypes.c_void_p)))
	hr  = fn(enumerator, data_flow, 1, ctypes.byref(col))
	return col if hr == 0 else None


def _collection_get_count(col):
	n = ctypes.c_uint(0)
	vt = ctypes.cast(col, ctypes.POINTER(ctypes.c_void_p))
	vt = ctypes.cast(vt[0], ctypes.POINTER(ctypes.c_void_p))
	fn = ctypes.cast(vt[VTBL_Collection_GetCount], ctypes.CFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p, ctypes.POINTER(ctypes.c_uint)))
	fn(col, ctypes.byref(n))
	return n.value


def _collection_item(col, i):
	dev = ctypes.c_void_p(0)
	vt  = ctypes.cast(col, ctypes.POINTER(ctypes.c_void_p))
	vt  = ctypes.cast(vt[0], ctypes.POINTER(ctypes.c_void_p))
	fn  = ctypes.cast(vt[VTBL_Collection_Item], ctypes.CFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p, ctypes.c_uint, ctypes.POINTER(ctypes.c_void_p)))
	hr  = fn(col, i, ctypes.byref(dev))
	return dev if hr == 0 else None


def _device_open_property_store(dev):
	ps = ctypes.c_void_p(0)
	vt = ctypes.cast(dev, ctypes.POINTER(ctypes.c_void_p))
	vt = ctypes.cast(vt[0], ctypes.POINTER(ctypes.c_void_p))
	fn = ctypes.cast(vt[VTBL_Device_OpenPropertyStore], ctypes.CFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p, ctypes.c_uint, ctypes.POINTER(ctypes.c_void_p)))
	hr = fn(dev, 0, ctypes.byref(ps))
	return ps if hr == 0 else None


class PROPVARIANT(ctypes.Structure):
	# Simplified: vt=VT_LPWSTR(31), padding, then pointer to wchar string
	_fields_ = [
		("vt",       ctypes.c_ushort),
		("pad1",     ctypes.c_ushort),
		("pad2",     ctypes.c_ushort),
		("pad3",     ctypes.c_ushort),
		("pwszVal",  ctypes.c_wchar_p),
	]


class PROPERTYKEY(ctypes.Structure):
	_fields_ = [("fmtid", ctypes.c_byte * 16), ("pid",   ctypes.c_ulong)]


def _get_device_friendly_name(dev):
	ps = _device_open_property_store(dev)
	if not ps:
		return ""
	try:
		key = PROPERTYKEY()
		key.fmtid = PKEY_FN_GUID
		key.pid   = PKEY_FN_PID
		pv = PROPVARIANT()
		vt = ctypes.cast(ps, ctypes.POINTER(ctypes.c_void_p))
		vt = ctypes.cast(vt[0], ctypes.POINTER(ctypes.c_void_p))
		fn = ctypes.cast(vt[VTBL_PropStore_GetValue], ctypes.CFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p, ctypes.POINTER(PROPERTYKEY), ctypes.POINTER(PROPVARIANT)))
		hr = fn(ps, ctypes.byref(key), ctypes.byref(pv))
		if hr == 0 and pv.vt == 31:   # VT_LPWSTR
			return pv.pwszVal or ""
		return ""
	except Exception:
		return ""
	finally:
		_release(ps)


def _get_device_id(dev):
	"""Return device ID string (if available)."""
	try:
		buf = ctypes.c_wchar_p()
		vt = ctypes.cast(dev, ctypes.POINTER(ctypes.c_void_p))
		vt = ctypes.cast(vt[0], ctypes.POINTER(ctypes.c_void_p))
		fn = ctypes.cast(vt[VTBL_Device_GetId], ctypes.CFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p, ctypes.POINTER(ctypes.c_wchar_p)))
		hr = fn(dev, ctypes.byref(buf))
		if hr == 0 and buf.value:
			val = buf.value
			try:
				ole32.CoTaskMemFree(buf)
			except Exception:
				pass
			return val
	except Exception:
		pass
	return ""


def _normalize_text(s: str) -> str:
	"""Normalize a device name: remove diacritics, lower-case and keep alnum+space."""
	if not s:
		return ""
	s = unicodedata.normalize('NFKD', s)
	s = ''.join(ch for ch in s if not unicodedata.combining(ch))
	s = s.lower()
	s = ''.join(ch if (ch.isalnum() or ch.isspace()) else ' ' for ch in s)
	parts = [p for p in s.split() if p]
	return ' '.join(parts)


def _device_activate_session_manager(dev):
	mgr = ctypes.c_void_p(0)
	vt  = ctypes.cast(dev, ctypes.POINTER(ctypes.c_void_p))
	vt  = ctypes.cast(vt[0], ctypes.POINTER(ctypes.c_void_p))
	fn  = ctypes.cast(vt[VTBL_Device_Activate], ctypes.CFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p, ctypes.c_byte * 16, ctypes.c_uint, ctypes.c_void_p, ctypes.POINTER(ctypes.c_void_p)))
	hr  = fn(dev, IID_IAudioSessionManager2, CLSCTX_ALL, None, ctypes.byref(mgr))
	return mgr if hr == 0 else None


def _get_session_enumerator(mgr):
	se = ctypes.c_void_p(0)
	vt = ctypes.cast(mgr, ctypes.POINTER(ctypes.c_void_p))
	vt = ctypes.cast(vt[0], ctypes.POINTER(ctypes.c_void_p))
	fn = ctypes.cast(vt[VTBL_ASM2_GetSessionEnumerator], ctypes.CFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p, ctypes.POINTER(ctypes.c_void_p)))
	hr = fn(mgr, ctypes.byref(se))
	return se if hr == 0 else None


def _session_enum_count(se):
	n  = ctypes.c_int(0)
	vt = ctypes.cast(se, ctypes.POINTER(ctypes.c_void_p))
	vt = ctypes.cast(vt[0], ctypes.POINTER(ctypes.c_void_p))
	fn = ctypes.cast(vt[VTBL_ASE_GetCount], ctypes.CFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p, ctypes.POINTER(ctypes.c_int)))
	fn(se, ctypes.byref(n))
	return n.value


def _session_enum_get(se, i):
	ctrl = ctypes.c_void_p(0)
	vt   = ctypes.cast(se, ctypes.POINTER(ctypes.c_void_p))
	vt   = ctypes.cast(vt[0], ctypes.POINTER(ctypes.c_void_p))
	fn   = ctypes.cast(vt[VTBL_ASE_GetSession], ctypes.CFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p, ctypes.c_int, ctypes.POINTER(ctypes.c_void_p)))
	hr   = fn(se, i, ctypes.byref(ctrl))
	return ctrl if hr == 0 else None


def _get_session_pid(ctrl):
	ctrl2 = _qi(ctrl, IID_IAudioSessionControl2)
	if not ctrl2:
		return 0
	try:
		pid = ctypes.c_uint(0)
		vt  = ctypes.cast(ctrl2, ctypes.POINTER(ctypes.c_void_p))
		vt  = ctypes.cast(vt[0], ctypes.POINTER(ctypes.c_void_p))
		fn  = ctypes.cast(vt[VTBL_ASC2_GetProcessId],
						  ctypes.CFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p,
										   ctypes.POINTER(ctypes.c_uint)))
		fn(ctrl2, ctypes.byref(pid))
		return pid.value
	finally:
		_release(ctrl2)


def _set_session_mute(ctrl, mute: bool):
	vol = _qi(ctrl, IID_ISimpleAudioVolume)
	if not vol:
		return False
	try:
		vt = ctypes.cast(vol, ctypes.POINTER(ctypes.c_void_p))
		vt = ctypes.cast(vt[0], ctypes.POINTER(ctypes.c_void_p))
		fn = ctypes.cast(vt[VTBL_SAV_SetMute],
						 ctypes.CFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p,
										  ctypes.c_bool, ctypes.c_void_p))
		hr = fn(vol, mute, None)
		return hr == 0
	finally:
		_release(vol)


def _get_session_mute(ctrl):
	vol = _qi(ctrl, IID_ISimpleAudioVolume)
	if not vol:
		return False
	try:
		m  = ctypes.c_bool(False)
		vt = ctypes.cast(vol, ctypes.POINTER(ctypes.c_void_p))
		vt = ctypes.cast(vt[0], ctypes.POINTER(ctypes.c_void_p))
		fn = ctypes.cast(vt[VTBL_SAV_GetMute],
						 ctypes.CFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p,
										  ctypes.POINTER(ctypes.c_bool)))
		fn(vol, ctypes.byref(m))
		return m.value
	finally:
		_release(vol)


def _is_sonar(name: str) -> bool:
	if not name:
		return False
	n = _normalize_text(name)
	if not n:
		return False
	for k in SONAR_KEYWORDS:
		if k in n:
			return True
	# heuristics: if 'sonar' appears, assume it's a Sonar virtual device
	if 'sonar' in n:
		return True
	# vendor + virtual hints
	if 'steel' in n and ('virtual' in n or 'audio' in n or 'microphone' in n or 'headset' in n):
		return True
	return False


def _is_target_device(name: str) -> bool:
	"""Return True if `name` should be targeted for muting.

	If `TARGET_DEVICE_KEYWORDS` is empty, fall back to Sonar detection.
	"""
	if not name:
		return False
	n = _normalize_text(name)
	if not n:
		return False
	if TARGET_DEVICE_KEYWORDS:
		for k in TARGET_DEVICE_KEYWORDS:
			k_norm = _normalize_text(k)
			if k_norm and k_norm in n:
				return True
		return False
	return _is_sonar(name)


def _discord_pids() -> set:
	pids = set()
	for p in psutil.process_iter(["pid", "name"]):
		try:
			if "discord" in p.info["name"].lower():
				pids.add(p.info["pid"])
		except Exception:
			pass
	return pids


def _iter_all_devices(data_flow: int = 2):
	"""Yield (device_ptr, friendly_name).

	`data_flow` follows IMMDeviceEnumerator::EnumAudioEndpoints: 0=render,1=capture,2=all
	Caller must _release each device.
	"""
	enum = _create_device_enumerator()
	if not enum:
		return
	try:
		col = _enum_audio_endpoints(enum, data_flow)
		if not col:
			return
		try:
			n = _collection_get_count(col)
			for i in range(n):
				dev = _collection_item(col, i)
				if dev:
					fname = _get_device_friendly_name(dev)
					yield dev, fname
		finally:
			_release(col)
	finally:
		_release(enum)


def mute_discord_in_sonar(mute: bool) -> int:
	pids = _discord_pids()
	if not pids:
		return 0
	total = 0
	# iterate only render (output) devices so we mute output sessions only
	for dev, fname in _iter_all_devices(data_flow=0):
		try:
			if not _is_target_device(fname):
				continue
			mgr = _device_activate_session_manager(dev)
			if not mgr:
				continue
			try:
				se = _get_session_enumerator(mgr)
				if not se:
					continue
				try:
					for i in range(_session_enum_count(se)):
						ctrl = _session_enum_get(se, i)
						if not ctrl:
							continue
						try:
							if _get_session_pid(ctrl) in pids:
								_set_session_mute(ctrl, mute)
								total += 1
						finally:
							_release(ctrl)
				finally:
					_release(se)
			finally:
				_release(mgr)
		finally:
			_release(dev)
	return total

def unmute_discord_all():
	mute_discord_in_sonar(False)

def run_diagnostics() -> list:
	lines = []
	ole32.CoInitialize(None)
	try:
		pids = _discord_pids()
		if pids:
			lines.append(f"[OK] Discord PIDs: {', '.join(str(p) for p in pids)}")
		else:
			lines.append("[!!] Discord not found running")
			return lines

		lines.append("--- All audio devices ---")
		sonar_devs = []
		enum = _create_device_enumerator()
		if not enum:
			lines.append("[!!] Failed to create device enumerator")
			return lines
		try:
			col = _enum_audio_endpoints(enum)
			if not col:
				lines.append("[!!] EnumAudioEndpoints failed")
				return lines
			try:
				n = _collection_get_count(col)
				lines.append(f"     Total devices: {n}")
				for i in range(n):
					dev = _collection_item(col, i)
					if not dev:
						continue
					try:
						fname = _get_device_friendly_name(dev)
						dev_id = _get_device_id(dev) or None
						is_s  = _is_sonar(fname)
						is_target = _is_target_device(fname)
						tag   = "[SONAR]" if is_s else "       "
						if dev_id:
							lines.append(f"{tag} {i} {fname} ({dev_id})")
						else:
							lines.append(f"{tag} {i} {fname}")
						if is_s:
							# keep Sonar devices for session listing, but mark target state
							sonar_devs.append((dev, fname, is_target))
						else:
							_release(dev)
					except Exception as ex:
						lines.append(f"       (error: {ex})")
						_release(dev)
			finally:
				_release(col)
		finally:
			_release(enum)

		if not sonar_devs:
			lines.append("[!!] No Sonar devices found - is SteelSeries Sonar app running?")
			return lines

		lines.append(f"[OK] {len(sonar_devs)} Sonar device(s) found")
		lines.append("--- Sessions on Sonar devices ---")

		for dev, fname, is_target in sonar_devs:
			try:
				mgr = _device_activate_session_manager(dev)
				if not mgr:
					lines.append(f"  {fname}: failed to activate session manager")
					continue
				try:
					se = _get_session_enumerator(mgr)
					if not se:
						lines.append(f"  {fname}: no session enumerator")
						continue
					try:
						sc = _session_enum_count(se)
						tmark = " [TARGET]" if is_target else ""
						lines.append(f"  {fname}:{tmark} {sc} session(s)")
						for i in range(sc):
							ctrl = _session_enum_get(se, i)
							if not ctrl:
								continue
							try:
								pid   = _get_session_pid(ctrl)
								is_dc = pid in pids
								try:    pname = psutil.Process(pid).name()
								except: pname = f"PID:{pid}"
								muted = _get_session_mute(ctrl)
								dc_tag = " <<DISCORD>>" if is_dc else ""
								m_tag  = " [MUTED]" if muted else " [active]"
								lines.append(f"    [{i}] {pname}{dc_tag}{m_tag}")
							finally:
								_release(ctrl)
					finally:
						_release(se)
				finally:
					_release(mgr)
			finally:
				_release(dev)

		lines.append("--- Mute attempt ---")
		n = mute_discord_in_sonar(True)
		if n > 0:
			lines.append(f"[OK] Muted {n} Discord session(s)!")
		else:
			lines.append("[!!] 0 sessions muted.")
			lines.append("     Make sure you are screen sharing WITH audio")
			lines.append("     and someone is watching with audio turned on,")
			lines.append("     then click Detect & Fix again.")
	finally:
		ole32.CoUninitialize()
	return lines


# ── Backend ───────────────────────────────────────────────────────────────────

class SonarFixApp:
	def __init__(self):
		self.mode            = "both"
		self.running         = False
		self.monitor_thread  = None
		self.status_callback = None
		self.log_callback    = None
		self._stop_event     = threading.Event()
		self.sessions_muted  = 0
		self.check_interval  = 1.0
		self.startup_enabled = self._check_startup()

	def set_mode(self, mode):     self.mode = mode
	def set_callbacks(self, status_cb=None, log_cb=None):
		self.status_callback = status_cb
		self.log_callback    = log_cb

	def _log(self, msg):
		if self.log_callback: self.log_callback(msg)
	def _set_status(self, status, count=0):
		if self.status_callback: self.status_callback(status, count)

	def start(self):
		if self.running: return
		self.running = True
		self._stop_event.clear()
		self.monitor_thread = threading.Thread(target=self._monitor_loop, daemon=True)
		self.monitor_thread.start()
		self._log("Protection started.")
		self._set_status("active", 0)

	def stop(self):
		if not self.running: return
		self.running = False
		self._stop_event.set()
		self._log("Stopping. Unmuting Discord...")
		unmute_discord_all()
		self._set_status("idle", 0)

	def _monitor_loop(self):
		ole32.CoInitialize(None)
		try:
			while not self._stop_event.wait(self.check_interval):
				try:
					count = mute_discord_in_sonar(True)
					if count != self.sessions_muted:
						self.sessions_muted = count
						self._set_status("active", count)
						if count > 0:
							self._log(f"Muted {count} Discord session(s) in Sonar devices.")
						else:
							self._log("Waiting for Discord screen share with audio...")
				except Exception as e:
					self._log(f"Monitor error: {e}")
		finally:
			ole32.CoUninitialize()

	def _check_startup(self):
		try:
			key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
				r"Software\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_READ)
			try:    winreg.QueryValueEx(key, APP_NAME); return True
			except: return False
			finally: winreg.CloseKey(key)
		except: return False

	def set_startup(self, enable):
		try:
			key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
				r"Software\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_SET_VALUE)
			if enable:
				path = (sys.executable if getattr(sys, "frozen", False)
						else f'"{sys.executable}" "{os.path.abspath(__file__)}"')
				winreg.SetValueEx(key, APP_NAME, 0, winreg.REG_SZ, path)
				self._log("Added to startup.")
			else:
				try:   winreg.DeleteValue(key, APP_NAME); self._log("Removed from startup.")
				except: pass
			winreg.CloseKey(key)
			self.startup_enabled = enable
		except Exception as e:
			self._log(f"Startup error: {e}")


# ── GUI ───────────────────────────────────────────────────────────────────────

import tkinter as tk
import tkinter.messagebox as msgbox


class AnimatedButton(tk.Canvas):
	def __init__(self, parent, text, command=None,
				 normal_color="#3a3d4a", hover_color="#4a4f60", press_color="#2a2d3a",
				 fg_color="#ffffff", font=("Segoe UI", 10, "bold"),
				 width=200, height=44, radius=8, **kw):
		super().__init__(parent, width=width, height=height,
						 bg=parent.cget("bg"), highlightthickness=0, **kw)
		self.command=command; self.text=text; self.font=font
		self.fg_color=fg_color; self.radius=radius
		self.btn_width=width; self.btn_height=height
		self.c_normal=normal_color; self.c_hover=hover_color; self.c_press=press_color
		self._cur=normal_color; self._anim_id=None
		self._redraw(normal_color)
		self.bind("<Enter>",           self._on_enter)
		self.bind("<Leave>",           self._on_leave)
		self.bind("<ButtonPress-1>",   self._on_press)
		self.bind("<ButtonRelease-1>", self._on_release)
		self.config(cursor="hand2")

	@staticmethod
	def _h2rgb(h):
		h=h.lstrip("#"); return int(h[0:2],16),int(h[2:4],16),int(h[4:6],16)
	@staticmethod
	def _rgb2h(r,g,b): return f"#{int(r):02x}{int(g):02x}{int(b):02x}"
	def _lerp(self,a,b,t):
		r1,g1,b1=self._h2rgb(a); r2,g2,b2=self._h2rgb(b)
		return self._rgb2h(r1+(r2-r1)*t,g1+(g2-g1)*t,b1+(b2-b1)*t)
	def _rrect(self,x1,y1,x2,y2,r,**kw):
		self.delete("bg")
		pts=[x1+r,y1,x2-r,y1,x2,y1,x2,y1+r,x2,y2-r,x2,y2,
			 x2-r,y2,x1+r,y2,x1,y2,x1,y2-r,x1,y1+r,x1,y1]
		self.create_polygon(pts,smooth=True,tags="bg",**kw)
	def _redraw(self,color):
		w,h,r=self.btn_width,self.btn_height,self.radius
		self._rrect(1,1,w-1,h-1,r,fill=color,outline="")
		self.delete("lbl")
		self.create_text(w//2,h//2,text=self.text,fill=self.fg_color,font=self.font,tags="lbl")
	def _animate_to(self,target,steps=8,delay=14):
		if self._anim_id: self.after_cancel(self._anim_id)
		self._step(self._cur,target,0,steps,delay)
	def _step(self,start,end,step,total,delay):
		t=min((step+1)/total,1.0); color=self._lerp(start,end,t)
		self._cur=color; self._redraw(color)
		if t<1.0: self._anim_id=self.after(delay,lambda:self._step(start,end,step+1,total,delay))
	def _on_enter(self,_):   self._animate_to(self.c_hover,6,12)
	def _on_leave(self,_):   self._animate_to(self.c_normal,8,12)
	def _on_press(self,_):   self._animate_to(self.c_press,3,8)
	def _on_release(self,_):
		self._animate_to(self.c_hover,4,10)
		if self.command: self.after(60,self.command)
	def resize(self,new_w):
		self.btn_width=new_w; self.config(width=new_w); self._redraw(self._cur)


class ModernApp(tk.Tk):
	BG="#202020"; SRF="#2c2c2c"; SRF2="#333333"; SRF3="#3d3d3d"; BORDER="#404040"
	ACCENT="#0078d4"; ACCENT_H="#1a88e4"; ACCENT_P="#005fa8"
	RED="#c42b1c"; RED_H="#d44030"; RED_P="#a02010"
	TEXT="#ffffff"; TEXT_S="#ababab"; TEXT_D="#737373"; GREEN="#6ccb5f"; TITLEBAR="#1c1c1c"

	def __init__(self):
		super().__init__()
		self.app=SonarFixApp()
		self.app.set_callbacks(status_cb=self._on_status,log_cb=self._on_log)
		self._setup_window(); self._build_chrome()
		self._show_mode_screen()
		self.after(500,self._auto_start)

	def _setup_window(self):
		self.title(APP_NAME); self.geometry("520x680"); self.minsize(480,620)
		self.configure(bg=self.BG); self.update_idletasks()
		x=(self.winfo_screenwidth()-520)//2; y=(self.winfo_screenheight()-680)//2
		self.geometry(f"+{x}+{y}"); self.protocol("WM_DELETE_WINDOW",self._on_close)

	def F(self,size,weight="normal",fam="Segoe UI"): return (fam,size,weight)

	def _build_chrome(self):
		tbar=tk.Frame(self,bg=self.TITLEBAR,height=50)
		tbar.pack(fill="x"); tbar.pack_propagate(False)
		inner=tk.Frame(tbar,bg=self.TITLEBAR)
		inner.pack(fill="both",expand=True,padx=18)
		dot=tk.Canvas(inner,width=14,height=14,bg=self.TITLEBAR,highlightthickness=0)
		dot.create_oval(0,0,14,14,fill=self.ACCENT,outline="")
		dot.pack(side="left",pady=18,padx=(0,10))
		tk.Label(inner,text="SteelSeries Sonar Echo Fix",
				 font=self.F(11,"bold"),bg=self.TITLEBAR,fg=self.TEXT).pack(side="left")
		tk.Label(inner,text=f"v{APP_VERSION}",
				 font=self.F(9),bg=self.TITLEBAR,fg=self.TEXT_D).pack(side="left",padx=8)
		tk.Frame(self,bg=self.BORDER,height=1).pack(fill="x")
		self.content=tk.Frame(self,bg=self.BG)
		self.content.pack(fill="both",expand=True,padx=24,pady=16)
		tk.Frame(self,bg=self.BORDER,height=1).pack(fill="x",side="bottom")
		bbar=tk.Frame(self,bg=self.TITLEBAR,height=30)
		bbar.pack(fill="x",side="bottom"); bbar.pack_propagate(False)
		self._sdot=tk.Canvas(bbar,width=8,height=8,bg=self.TITLEBAR,highlightthickness=0)
		self._sdot.pack(side="left",padx=(14,5),pady=11); self._dot_color(self.TEXT_D)
		self._slbl=tk.Label(bbar,text="Idle  protection not active",
							font=self.F(8),bg=self.TITLEBAR,fg=self.TEXT_D)
		self._slbl.pack(side="left")
		self._scnt=tk.Label(bbar,text="",font=self.F(8,"bold"),bg=self.TITLEBAR,fg=self.GREEN)
		self._scnt.pack(side="right",padx=14)
		tk.Label(bbar,text="Development by CodeXart Studio's",
				 font=self.F(7),bg=self.TITLEBAR,fg="#505050").pack(side="right",padx=(0,6))

	def _dot_color(self,c):
		self._sdot.delete("all"); self._sdot.create_oval(0,0,8,8,fill=c,outline="")
	def _clear(self):
		for w in self.content.winfo_children(): w.destroy()

	def _show_mode_screen(self):
		self._clear()
		self.mode_var = tk.StringVar(value="both")
		self._cards = {}
		tk.Label(self.content, text="Select Your Setup",
				 font=self.F(17,"bold"), bg=self.BG, fg=self.TEXT).pack(pady=(4,2))
		tk.Label(self.content,
				 text="Choose which SteelSeries Sonar devices you are using.\n"
					  "The app will silently prevent Discord echo on screen share.",
				 font=self.F(9), bg=self.BG, fg=self.TEXT_S, justify="center").pack(pady=(0,18))
		for lbl, val, desc in [
			("Microphone  -  SteelSeries Sonar Microphone", "microphone", "Sonar - Microphone + Chat devices"),
			("Headset  -  SteelSeries Sonar Headset",       "headset",    "Sonar - Gaming + Media devices"),
			("Both  -  Microphone and Headset",             "both",       "Recommended: covers all Sonar virtual devices"),
		]:
			self._mode_card(lbl, val, desc)
		self.mode_var.trace_add("write", lambda *_: self._refresh_cards())
		self.after(50, self._refresh_cards)
		tk.Frame(self.content, bg=self.BORDER, height=1).pack(fill="x", pady=14)
		sf = tk.Frame(self.content, bg=self.SRF, padx=16, pady=10)
		sf.pack(fill="x")
		tk.Label(sf, text="Launch on Windows startup",
				 font=self.F(10), bg=self.SRF, fg=self.TEXT).pack(side="left")
		self.startup_var = tk.BooleanVar(value=self.app.startup_enabled)
		tk.Checkbutton(sf, variable=self.startup_var, bg=self.SRF, fg=self.ACCENT,
					   selectcolor=self.SRF, activebackground=self.SRF, relief="flat", bd=0,
					   command=lambda: self.app.set_startup(self.startup_var.get())).pack(side="right")
		tk.Frame(self.content, bg=self.BG, height=14).pack()
		bf = tk.Frame(self.content, bg=self.BG)
		bf.pack(fill="x")
		self._start_btn = AnimatedButton(bf, text="START PROTECTION", command=self._start,
			normal_color=self.ACCENT, hover_color=self.ACCENT_H, press_color=self.ACCENT_P,
			fg_color="#ffffff", font=self.F(11,"bold"), width=470, height=46, radius=8)
		self._start_btn.pack(fill="x", expand=True)
		bf.bind("<Configure>", lambda e: self._start_btn.resize(max(e.width-2, 100)))

	def _mode_card(self,label,val,desc):
		outer=tk.Frame(self.content,bg=self.BG); outer.pack(fill="x",pady=3)
		frame=tk.Frame(outer,bg=self.SRF2,padx=14,pady=11,cursor="hand2",
					   highlightthickness=1,highlightbackground=self.BORDER); frame.pack(fill="x")
		bar=tk.Frame(frame,bg=self.SRF2,width=3); bar.pack(side="left",fill="y",padx=(0,12))
		rb=tk.Radiobutton(frame,variable=self.mode_var,value=val,
						  bg=self.SRF2,fg=self.ACCENT,selectcolor=self.SRF2,
						  activebackground=self.SRF2,relief="flat",bd=0); rb.pack(side="left")
		inner=tk.Frame(frame,bg=self.SRF2); inner.pack(side="left",fill="x",expand=True,padx=8)
		tk.Label(inner,text=label,font=self.F(10,"bold"),bg=self.SRF2,fg=self.TEXT,anchor="w").pack(fill="x")
		tk.Label(inner,text=desc, font=self.F(9),       bg=self.SRF2,fg=self.TEXT_D,anchor="w").pack(fill="x")
		self._cards[val]=(frame,bar)
		all_w=[frame,inner,bar,rb]+list(inner.winfo_children())
		def hover_all(ws,on):
			c=self.SRF3 if on else self.SRF2
			for w in ws:
				try: w.config(bg=c)
				except: pass
				for ch in w.winfo_children():
					try: ch.config(bg=c)
					except: pass
		for w in all_w:
			w.bind("<Enter>",   lambda e,ws=all_w:hover_all(ws,True))
			w.bind("<Leave>",   lambda e,ws=all_w:hover_all(ws,False))
			w.bind("<Button-1>",lambda e,v=val:self.mode_var.set(v))

	def _refresh_cards(self):
		sel=self.mode_var.get()
		for val,(frame,bar) in self._cards.items():
			if val==sel: frame.config(highlightbackground=self.ACCENT); bar.config(bg=self.ACCENT)
			else:        frame.config(highlightbackground=self.BORDER);  bar.config(bg=self.SRF2)

	def _start(self):
		self.app.set_mode(self.mode_var.get()); self.app.start(); self._show_active_screen()

	def _show_active_screen(self):
		self._clear()
		self._pc=tk.Canvas(self.content,width=110,height=110,bg=self.BG,highlightthickness=0)
		self._pc.pack(pady=(8,4)); self._ring_step=0; self._anim_rings()
		tk.Label(self.content,text="PROTECTION ACTIVE",
				 font=self.F(15,"bold"),bg=self.BG,fg=self.GREEN).pack()
		mt={"microphone":"Microphone only","headset":"Headset only","both":"Microphone and Headset"}
		tk.Label(self.content,text=f"Mode: {mt.get(self.app.mode,self.app.mode)}",
				 font=self.F(9),bg=self.BG,fg=self.TEXT_D).pack(pady=(2,10))
		tk.Frame(self.content,bg=self.BORDER,height=1).pack(fill="x",pady=4)
		info=tk.Frame(self.content,bg=self.SRF,padx=14,pady=10); info.pack(fill="x",pady=8)
		for icon,line in [("+","Monitoring Sonar virtual devices every second"),
						  ("+","Auto-mutes Discord in Sonar Microphone, Chat, Gaming devices"),
						  ("+","You still hear Discord and your mic works normally"),
						  ("+","Others will not hear echoes or feedback from you")]:
			row=tk.Frame(info,bg=self.SRF); row.pack(fill="x",pady=2)
			tk.Label(row,text=icon,font=self.F(9,"bold"),bg=self.SRF,fg=self.GREEN,width=2).pack(side="left")
			tk.Label(row,text=line,font=self.F(9),       bg=self.SRF,fg=self.TEXT_S,anchor="w").pack(side="left")
		tk.Frame(self.content,bg=self.BORDER,height=1).pack(fill="x",pady=8)
		tk.Label(self.content,text="Activity Log",font=self.F(9,"bold"),
				 bg=self.BG,fg=self.TEXT_D,anchor="w").pack(fill="x")
		log_outer=tk.Frame(self.content,bg=self.SRF,padx=8,pady=8,
						   highlightthickness=1,highlightbackground=self.BORDER)
		log_outer.pack(fill="both",expand=True,pady=(4,8))
		self.log_text=tk.Text(log_outer,bg=self.SRF,fg=self.TEXT_D,font=("Consolas",9),
							  relief="flat",bd=0,state="disabled",wrap="word",height=6)
		self.log_text.pack(fill="both",expand=True)
		df=tk.Frame(self.content,bg=self.BG); df.pack(fill="x",pady=(0,5))
		self._detect_btn=AnimatedButton(df,text="DETECT & FIX NOW",command=self._manual_detect,
			normal_color=self.SRF2,hover_color=self.SRF3,press_color="#252525",
			fg_color=self.ACCENT,font=self.F(10,"bold"),width=470,height=38,radius=8)
		self._detect_btn.pack(fill="x",expand=True)
		df.bind("<Configure>",lambda e:self._detect_btn.resize(max(e.width-2,100)))
		sf=tk.Frame(self.content,bg=self.BG); sf.pack(fill="x")
		self._stop_btn=AnimatedButton(sf,text="STOP PROTECTION",command=self._stop,
			normal_color=self.RED,hover_color=self.RED_H,press_color=self.RED_P,
			fg_color="#ffffff",font=self.F(11,"bold"),width=470,height=46,radius=8)
		self._stop_btn.pack(fill="x",expand=True)
		sf.bind("<Configure>",lambda e:self._stop_btn.resize(max(e.width-2,100)))

	def _anim_rings(self):
		if not self.app.running: return
		try:
			c=self._pc; cx,cy=55,55; c.delete("ring")
			t=(self._ring_step%40)/40.0; r=22+t*32
			c.create_oval(cx-r,cy-r,cx+r,cy+r,outline=self.GREEN,width=2,tags="ring")
			t2=((self._ring_step+20)%40)/40.0; r2=22+t2*32
			c.create_oval(cx-r2,cy-r2,cx+r2,cy+r2,outline=self.GREEN,width=1,tags="ring")
			c.delete("center"); c.create_oval(cx-14,cy-14,cx+14,cy+14,fill=self.GREEN,outline="",tags="center")
			self._ring_step+=1; self.after(40,self._anim_rings)
		except: pass

	def _stop(self): self.app.stop(); self._show_mode_screen()

	def _manual_detect(self):
		self._on_log("--- Manual Detect & Fix ---")
		def run():
			try:
				for line in run_diagnostics(): self._on_log(line)
			except Exception as e: self._on_log(f"Error: {e}")
		threading.Thread(target=run,daemon=True).start()

	def _on_status(self,status,count):
		try:
			if status=="active":
				self._dot_color(self.GREEN)
				if count>0:
					self._slbl.config(text=f"Active  {count} session(s) muted",fg=self.GREEN)
					self._scnt.config(text=f"{count} muted")
				else:
					self._slbl.config(text="Active  waiting for Discord screen share",fg=self.TEXT_D)
					self._scnt.config(text="")
			else:
				self._dot_color(self.TEXT_D)
				self._slbl.config(text="Idle  protection not active",fg=self.TEXT_D)
				self._scnt.config(text="")
		except: pass

	def _on_log(self,msg):
		try:
			if hasattr(self,"log_text"):
				self.log_text.config(state="normal")
				self.log_text.insert("end",f"> {msg}\n")
				self.log_text.see("end")
				self.log_text.config(state="disabled")
		except: pass

	def _auto_start(self):
		if "--autostart" in sys.argv:
			self.app.set_mode("both"); self.app.start(); self._show_active_screen()

	def _on_close(self):
		"""Called when X button pressed — ask user what to do."""
		if self.app.running:
			# Custom dialog: minimize to tray OR stop and quit
			choice = _ask_close_action(self)
			if choice == "tray":
				self._minimize_to_tray()
			elif choice == "quit":
				self.app.stop()
				self._tray_stop()
				self.destroy()
			# else: cancelled, do nothing
		else:
			choice = _ask_close_action(self)
			if choice == "tray":
				self._minimize_to_tray()
			elif choice == "quit":
				self._tray_stop()
				self.destroy()

	def _minimize_to_tray(self):
		"""Hide window and show system tray icon."""
		self.withdraw()
		if not hasattr(self, "_tray_icon") or self._tray_icon is None:
			self._tray_icon = _create_tray_icon(self)
			t = threading.Thread(target=self._tray_icon.run, daemon=True)
			t.start()

	def _restore_from_tray(self):
		"""Show window again from tray."""
		self.deiconify()
		self.lift()
		self.focus_force()

	def _tray_stop(self):
		"""Stop tray icon if running."""
		try:
			if hasattr(self, "_tray_icon") and self._tray_icon:
				self._tray_icon.stop()
				self._tray_icon = None
		except Exception:
			pass


def _ask_close_action(parent):
	"""
	Custom modal dialog.
	Returns: 'tray', 'quit', or None (cancelled).
	"""
	result = [None]
	dlg = tk.Toplevel(parent)
	dlg.title("Close Application")
	dlg.configure(bg="#202020")
	dlg.resizable(False, False)
	dlg.grab_set()

	# Center on parent
	parent.update_idletasks()
	px = parent.winfo_x() + parent.winfo_width()  // 2
	py = parent.winfo_y() + parent.winfo_height() // 2
	dlg.geometry(f"380x200+{px-190}+{py-100}")

	tk.Label(dlg, text="What would you like to do?",
			 font=("Segoe UI", 12, "bold"), bg="#202020", fg="#ffffff").pack(pady=(20,6))
	tk.Label(dlg, text="Protection is still active." if parent.app.running else "Choose an option.",
			 font=("Segoe UI", 9), bg="#202020", fg="#ababab").pack(pady=(0,16))

	btn_frame = tk.Frame(dlg, bg="#202020")
	btn_frame.pack(fill="x", padx=20)

	def do_tray():
		result[0] = "tray"; dlg.destroy()
	def do_quit():
		result[0] = "quit"; dlg.destroy()
	def do_cancel():
		result[0] = None;  dlg.destroy()

	# Tray button
	tb = tk.Button(btn_frame, text="Run in Background",
				   font=("Segoe UI", 10, "bold"), bg="#0078d4", fg="white",
				   activebackground="#1a88e4", activeforeground="white",
				   relief="flat", bd=0, padx=16, pady=10,
				   cursor="hand2", command=do_tray)
	tb.pack(side="left", expand=True, fill="x", padx=(0,6))

	# Quit button
	qb = tk.Button(btn_frame, text="Stop & Quit",
				   font=("Segoe UI", 10), bg="#333333", fg="#ffffff",
				   activebackground="#3d3d3d", activeforeground="white",
				   relief="flat", bd=0, padx=16, pady=10,
				   cursor="hand2", command=do_quit)
	qb.pack(side="left", expand=True, fill="x", padx=(0,6))

	# Cancel button
	cb = tk.Button(btn_frame, text="Cancel",
				   font=("Segoe UI", 10), bg="#2c2c2c", fg="#ababab",
				   activebackground="#3d3d3d", activeforeground="white",
				   relief="flat", bd=0, padx=16, pady=10,
				   cursor="hand2", command=do_cancel)
	cb.pack(side="left", expand=True, fill="x")

	dlg.protocol("WM_DELETE_WINDOW", do_cancel)
	parent.wait_window(dlg)
	return result[0]


def _make_tray_image(active=True):
	"""Create a simple icon image for the system tray."""
	size = 64
	img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
	draw = ImageDraw.Draw(img)
	color = (108, 203, 95, 255) if active else (115, 115, 115, 255)  # green or grey
	draw.ellipse([8, 8, 56, 56], fill=color)
	# Small inner dot
	draw.ellipse([24, 24, 40, 40], fill=(255, 255, 255, 200))
	return img


def _create_tray_icon(app_window):
	"""Create and return a pystray Icon."""
	img = _make_tray_image(active=app_window.app.running)

	def on_open(icon, item):
		app_window.after(0, app_window._restore_from_tray)

	def on_quit(icon, item):
		app_window.app.stop()
		icon.stop()
		app_window.after(0, app_window.destroy)

	menu = pystray.Menu(
		pystray.MenuItem("Open", on_open, default=True),
		pystray.Menu.SEPARATOR,
		pystray.MenuItem("Stop & Quit", on_quit),
	)

	icon = pystray.Icon(
		"SonarEchoFix",
		img,
		"Sonar Echo Fix — Active" if app_window.app.running else "Sonar Echo Fix",
		menu
	)
	return icon


def main():
	if sys.platform != "win32":
		print("This application only works on Windows."); sys.exit(1)
	# Quick diagnostic CLI mode
	if "--diag" in sys.argv:
		ole32.CoInitialize(None)
		try:
			for line in run_diagnostics():
				print(line)
		finally:
			ole32.CoUninitialize()
		return
	ModernApp().mainloop()


if __name__ == "__main__":
	main()
