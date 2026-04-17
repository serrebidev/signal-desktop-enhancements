import ctypes
import ctypes.wintypes
import time
from collections import deque

import appModuleHandler
import scriptHandler
import tones
import ui
from logHandler import log

import comtypes.client
from comtypes import POINTER, HRESULT


comtypes.client.GetModule("oleacc.dll")
from comtypes.gen.Accessibility import IAccessible


CHILDID_SELF = 0
OBJID_CLIENT = 0xFFFFFFFC
SELFLAG_TAKEFOCUS = 0x1
SELFLAG_TAKESELECTION = 0x2

ROLE_DOCUMENT = "document"
ROLE_PROPERTY_PAGE = "property page"
ROLE_PAGE_TAB = "page tab"
ROLE_PUSH_BUTTON = "push button"
ROLE_MENU_BUTTON = "menu button"
ROLE_EDITABLE_TEXT = "editable text"
ROLE_TABLE = "table"
ROLE_LIST = "list"
ROLE_LIST_ITEM = "list item"
ROLE_ROW = "row"
ROLE_GROUPING = "grouping"

FOCUSABLE = "focusable"
FOCUSED = "focused"
SELECTED = "selected"
OFFSCREEN = "offscreen"

TAB_CONTACTS = ("Contacts", "Chats")
TAB_CALLS = ("Calls",)
TAB_STORIES = ("Stories",)
TAB_SETTINGS = ("Settings", "Preferences")

SETTINGS_NAV_BUTTONS = frozenset(
	name.lower() for name in (
		"General",
		"Appearance",
		"Chats",
		"Calls",
		"Notifications",
		"Privacy",
		"Data usage",
		"Backups",
		"Donate to Signal",
	)
)
SETTINGS_SKIP_BUTTONS = frozenset(
	name.lower() for name in (
		"Hide Tabs",
		"Open username link screen",
		"Edit photo",
	)
)

MAX_DEPTH = 18
ROOT_TIMEOUT = 0.8
PAGE_TIMEOUT = 0.8

user32 = ctypes.windll.user32
oleacc = ctypes.windll.oleacc

oleacc.AccessibleObjectFromWindow.restype = HRESULT
oleacc.GetRoleTextW.argtypes = [
	ctypes.wintypes.DWORD,
	ctypes.wintypes.LPWSTR,
	ctypes.wintypes.UINT,
]
oleacc.GetRoleTextW.restype = ctypes.wintypes.UINT
oleacc.GetStateTextW.argtypes = [
	ctypes.wintypes.DWORD,
	ctypes.wintypes.LPWSTR,
	ctypes.wintypes.UINT,
]
oleacc.GetStateTextW.restype = ctypes.wintypes.UINT


def _normalize(text):
	return " ".join((text or "").strip().lower().split())


def _role_text(value):
	buffer = ctypes.create_unicode_buffer(128)
	length = oleacc.GetRoleTextW(int(value), buffer, len(buffer))
	return buffer.value if length else str(value)


def _state_texts(mask):
	states = set()
	bit = 1
	while bit <= 0x40000000:
		if mask & bit:
			buffer = ctypes.create_unicode_buffer(128)
			length = oleacc.GetStateTextW(bit, buffer, len(buffer))
			if length:
				states.add(buffer.value.lower())
		bit <<= 1
	return states


def _acc_identity(acc):
	try:
		return int(ctypes.cast(acc, ctypes.c_void_p).value or 0)
	except Exception:
		return id(acc)


def _name(acc):
	try:
		return acc.accName(CHILDID_SELF) or ""
	except Exception:
		return ""


def _role(acc):
	try:
		return _role_text(acc.accRole(CHILDID_SELF)).lower()
	except Exception:
		return ""


def _states(acc):
	try:
		return _state_texts(int(acc.accState(CHILDID_SELF)))
	except Exception:
		return set()


def _is_focusable(acc):
	return FOCUSABLE in _states(acc)


def _is_visible(acc):
	return OFFSCREEN not in _states(acc)


def _children(acc):
	try:
		count = int(acc.accChildCount)
	except Exception:
		count = 0
	for index in range(1, count + 1):
		try:
			child = acc.accChild(index)
		except Exception:
			continue
		if isinstance(child, int):
			continue
		try:
			yield child.QueryInterface(IAccessible)
		except Exception:
			continue


def _iter_tree(root, max_depth=MAX_DEPTH, timeout=ROOT_TIMEOUT):
	if root is None:
		return
	start = time.time()
	queue = deque([(root, 0, ())])
	seen = set()
	while queue:
		if time.time() - start > timeout:
			return
		acc, depth, chain = queue.popleft()
		identity = _acc_identity(acc)
		if identity in seen:
			continue
		seen.add(identity)
		yield acc, depth, chain
		if depth >= max_depth:
			continue
		child_chain = chain + (acc,)
		for child in _children(acc):
			queue.append((child, depth + 1, child_chain))


def _find_first(root, predicate, max_depth=MAX_DEPTH, timeout=ROOT_TIMEOUT):
	for acc, depth, chain in _iter_tree(root, max_depth=max_depth, timeout=timeout):
		try:
			if predicate(acc, depth, chain):
				return acc, chain
		except Exception:
			continue
	return None, ()


def _subtree_has_name(root, names, max_depth=4):
	patterns = tuple(_normalize(name) for name in names)
	for acc, _depth, _chain in _iter_tree(root, max_depth=max_depth, timeout=0.25):
		name = _normalize(_name(acc))
		if not name:
			continue
		for pattern in patterns:
			if pattern in name:
				return True
	return False


def _find_focusable_descendant(root, names=(), roles=(), max_depth=6):
	role_names = tuple(role.lower() for role in roles)
	name_patterns = tuple(_normalize(name) for name in names)
	for acc, _depth, _chain in _iter_tree(root, max_depth=max_depth, timeout=0.35):
		if not _is_focusable(acc) or not _is_visible(acc):
			continue
		role = _role(acc)
		name = _normalize(_name(acc))
		if role_names and role not in role_names:
			continue
		if name_patterns and not any(pattern in name for pattern in name_patterns):
			continue
		return acc
	return None


def _do_default_action(acc):
	try:
		acc.accDoDefaultAction(CHILDID_SELF)
		return True
	except Exception:
		return False


def _focus_acc(acc):
	try:
		acc.accSelect(SELFLAG_TAKEFOCUS, CHILDID_SELF)
		return True
	except Exception:
		return False


def _activate_tab(acc):
	if _do_default_action(acc):
		return True
	try:
		acc.accSelect(SELFLAG_TAKEFOCUS | SELFLAG_TAKESELECTION, CHILDID_SELF)
		return True
	except Exception:
		return False


def _announce_found(label):
	ui.message(label)


def _announce_missing(label):
	ui.message("%s not found" % label)
	tones.beep(150, 60)


def _foreground_root_hwnd():
	hwnd = user32.GetForegroundWindow()
	if not hwnd:
		return 0
	GA_ROOT = 2
	try:
		hwnd = user32.GetAncestor(hwnd, GA_ROOT) or hwnd
	except Exception:
		pass
	return hwnd


def _render_hwnd(root_hwnd):
	children = []
	enum_proc = ctypes.WINFUNCTYPE(
		ctypes.wintypes.BOOL,
		ctypes.wintypes.HWND,
		ctypes.wintypes.LPARAM,
	)

	def _enum_child(hwnd, lparam):
		class_name = ctypes.create_unicode_buffer(256)
		user32.GetClassNameW(hwnd, class_name, len(class_name))
		children.append((hwnd, class_name.value))
		return True

	user32.EnumChildWindows(root_hwnd, enum_proc(_enum_child), 0)
	for hwnd, class_name in children:
		if class_name == "Chrome_RenderWidgetHostHWND":
			return hwnd
	return 0


def _signal_root():
	root_hwnd = _foreground_root_hwnd()
	if not root_hwnd:
		return None
	render_hwnd = _render_hwnd(root_hwnd)
	if not render_hwnd:
		return None
	ptr = ctypes.c_void_p()
	result = oleacc.AccessibleObjectFromWindow(
		render_hwnd,
		OBJID_CLIENT,
		ctypes.byref(IAccessible._iid_),
		ctypes.byref(ptr),
	)
	if result != 0 or not ptr.value:
		return None
	return ctypes.cast(ptr, POINTER(IAccessible))


def _find_tab(root, names):
	expected = {_normalize(name) for name in names}
	return _find_first(
		root,
		lambda acc, _depth, _chain: (
			_role(acc) == ROLE_PAGE_TAB
			and _normalize(_name(acc)) in expected
			and OFFSCREEN not in _states(acc)
		),
		max_depth=12,
		timeout=0.35,
	)


def _selected_tab_name(root):
	selected, _chain = _find_first(
		root,
		lambda acc, _depth, _chain: (
			_role(acc) == ROLE_PAGE_TAB
			and SELECTED in _states(acc)
			and OFFSCREEN not in _states(acc)
		),
		max_depth=12,
		timeout=0.35,
	)
	return _name(selected) if selected is not None else ""


def _find_page(root, page_name):
	expected = _normalize(page_name)
	return _find_first(
		root,
		lambda acc, _depth, _chain: (
			_role(acc) == ROLE_PROPERTY_PAGE
			and _normalize(_name(acc)) == expected
			and OFFSCREEN not in _states(acc)
		),
		max_depth=12,
		timeout=0.35,
	)


def _wait_for_page(page_name, timeout=PAGE_TIMEOUT):
	end = time.time() + timeout
	root = _signal_root()
	page = None
	while time.time() < end:
		root = _signal_root()
		page, _chain = _find_page(root, page_name)
		if page is not None:
			return root, page
		time.sleep(0.05)
	page, _chain = _find_page(root, page_name)
	return root, page


def _wait_for_selected_tab(tab_names, timeout=PAGE_TIMEOUT):
	expected = {_normalize(name) for name in tab_names}
	end = time.time() + timeout
	root = _signal_root()
	while time.time() < end:
		root = _signal_root()
		if root is not None and _normalize(_selected_tab_name(root)) in expected:
			return root
		time.sleep(0.05)
	return _signal_root()


def _activate_page(page_name, tab_names):
	root = _signal_root()
	if root is None:
		return None, None
	expected = {_normalize(name) for name in tab_names}
	page, _chain = _find_page(root, page_name)
	if _normalize(_selected_tab_name(root)) in expected:
		if page is not None:
			return root, page
		return _wait_for_page(page_name)
	tab, _chain = _find_tab(root, tab_names)
	if tab is not None:
		_activate_tab(tab)
		time.sleep(0.15)
	root = _wait_for_selected_tab(tab_names)
	if root is None:
		return None, None
	page_root, page = _wait_for_page(page_name)
	return page_root or root, page


def _focus_chats_list(page):
	grid, _chain = _find_first(
		page,
		lambda acc, _depth, _chain: (
			_role(acc) == ROLE_TABLE
			and _normalize(_name(acc)) == "grid"
			and _is_visible(acc)
		),
		max_depth=10,
		timeout=0.25,
	)
	if grid is not None:
		return _focus_acc(grid)
	return False


def _focus_calls_list(page):
	fallback = None
	for acc, _depth, chain in _iter_tree(page, max_depth=16, timeout=0.45):
		if (
			_role(acc) != ROLE_TABLE
			or _normalize(_name(acc)) != "grid"
			or not _is_visible(acc)
		):
			continue
		if fallback is None:
			fallback = acc
		for ancestor in reversed(chain):
			if _subtree_has_name(ancestor, ("New Call",), max_depth=4):
				return _focus_acc(acc)
	if fallback is not None:
		return _focus_acc(fallback)
	return False


def _focus_message_field(page):
	anchor, chain = _find_first(
		page,
		lambda acc, _depth, _chain: (
			_role(acc) == ROLE_PUSH_BUTTON
			and "emoji" in _normalize(_name(acc))
			and _is_visible(acc)
		),
		max_depth=16,
		timeout=0.35,
	)
	if anchor is None or not chain:
		return False
	parent = chain[-1]
	children = list(_children(parent))
	best = None
	best_score = (-1, -1)
	for child in children:
		if _acc_identity(child) == _acc_identity(anchor):
			continue
		for acc, depth, _chain in _iter_tree(child, max_depth=8, timeout=0.2):
			if (
				_role(acc) != ROLE_GROUPING
				or not _is_focusable(acc)
				or not _is_visible(acc)
			):
				continue
			hint = 1 if _subtree_has_name(
				acc,
				("Message", "Type a message", "Write a message"),
				max_depth=4,
			) else 0
			score = (hint, depth)
			if score > best_score:
				best = acc
				best_score = score
	if best is not None:
		return _focus_acc(best)
	fallback = _find_focusable_descendant(
		parent,
		names=("Message", "Type a message", "Write a message"),
		roles=(ROLE_GROUPING, ROLE_EDITABLE_TEXT),
		max_depth=10,
	)
	if fallback is not None:
		return _focus_acc(fallback)
	return False


def _focus_chat_history(page):
	history_list, _chain = _find_first(
		page,
		lambda acc, _depth, _chain: (
			_role(acc) == ROLE_LIST
			and _is_visible(acc)
			and _subtree_has_name(
				acc,
				("Reply to Message", "React to Message", "More actions"),
				max_depth=4,
			)
		),
		max_depth=18,
		timeout=0.5,
	)
	if history_list is None:
		return False
	fallback = None
	for acc, _depth, _chain in _iter_tree(history_list, max_depth=6, timeout=0.25):
		if _role(acc) != ROLE_ROW or not _is_focusable(acc) or not _is_visible(acc):
			continue
		fallback = acc
		if FOCUSED in _states(acc):
			return _focus_acc(acc)
	if fallback is not None:
		return _focus_acc(fallback)
	return False


def _focus_stories(page):
	story_buttons = []
	search = None
	context_menu = None
	for acc, _depth, _chain in _iter_tree(page, max_depth=12, timeout=0.35):
		role = _role(acc)
		name = _normalize(_name(acc))
		if not _is_visible(acc):
			continue
		if role == ROLE_EDITABLE_TEXT and name == "search":
			search = acc
		elif role == ROLE_PUSH_BUTTON and name == "add a story":
			story_buttons.append(acc)
		elif role == ROLE_PUSH_BUTTON and name == "context menu":
			context_menu = acc
	if len(story_buttons) >= 2:
		return _focus_acc(story_buttons[1])
	if story_buttons:
		return _focus_acc(story_buttons[0])
	if context_menu is not None:
		return _focus_acc(context_menu)
	if search is not None:
		return _focus_acc(search)
	return False


def _focus_settings(page):
	fallback = None
	for acc, _depth, chain in _iter_tree(page, max_depth=14, timeout=0.45):
		if (
			_role(acc) != ROLE_PUSH_BUTTON
			or not _is_focusable(acc)
			or not _is_visible(acc)
		):
			continue
		name = _normalize(_name(acc))
		if not name or name in SETTINGS_NAV_BUTTONS or name in SETTINGS_SKIP_BUTTONS:
			continue
		if fallback is None:
			fallback = acc
		for ancestor in reversed(chain):
			if _subtree_has_name(ancestor, ("Profile",), max_depth=3):
				return _focus_acc(acc)
	if fallback is not None:
		return _focus_acc(fallback)
	return False


def _dump_tree(root, max_depth=14):
	lines = ["=== Signal Accessibility Tree ===", ""]
	count = 0
	for acc, depth, _chain in _iter_tree(root, max_depth=max_depth, timeout=1.5):
		role = _role(acc)
		name = _name(acc)
		states = ", ".join(sorted(_states(acc)))
		if depth == 0 or name or states:
			lines.append(
				"%s[%s] name=%r states=%r"
				% ("  " * min(depth, 10), role, name, states)
			)
			count += 1
		if count >= 400:
			lines.append("... (stopped at 400 items)")
			break
	lines.append("")
	lines.append("=== %d items logged ===" % count)
	return "\n".join(lines)


class AppModule(appModuleHandler.AppModule):
	scriptCategory = "Signal Desktop Enhancements"

	@scriptHandler.script(
		gesture="kb:alt+1",
		description="Focus the chats list",
		category="Signal Desktop Enhancements",
	)
	def script_focusChatsShortcut(self, gesture):
		self.script_focusChatList(gesture)

	@scriptHandler.script(
		gesture="kb:alt+2",
		description="Focus the calls list",
		category="Signal Desktop Enhancements",
	)
	def script_focusCallsShortcut(self, gesture):
		root, page = _activate_page("Calls", TAB_CALLS)
		scope = page or root
		if scope is not None and _focus_calls_list(scope):
			_announce_found("Calls list")
			return
		_announce_missing("Calls list")

	@scriptHandler.script(
		gesture="kb:alt+3",
		description="Focus stories",
		category="Signal Desktop Enhancements",
	)
	def script_focusStoriesShortcut(self, gesture):
		root, page = _activate_page("Stories", TAB_STORIES)
		scope = page or root
		if scope is not None and _focus_stories(scope):
			_announce_found("Stories")
			return
		_announce_missing("Stories")

	@scriptHandler.script(
		gesture="kb:alt+4",
		description="Focus the main settings pane",
		category="Signal Desktop Enhancements",
	)
	def script_focusSettingsShortcut(self, gesture):
		root, page = _activate_page("Settings", TAB_SETTINGS)
		scope = page or root
		if scope is not None and _focus_settings(scope):
			_announce_found("Settings")
			return
		_announce_missing("Settings")

	@scriptHandler.script(
		gesture="kb:control+shift+a",
		description="Focus the list of existing chats",
		category="Signal Desktop Enhancements",
	)
	def script_focusChatList(self, gesture):
		root, page = _activate_page("Chats", TAB_CONTACTS)
		scope = page or root
		if scope is not None and _focus_chats_list(scope):
			_announce_found("Chats list")
			return
		_announce_missing("Chats list")

	@scriptHandler.script(
		gesture="kb:control+1",
		description="Focus the message field",
		category="Signal Desktop Enhancements",
	)
	def script_focusMessageField(self, gesture):
		root, page = _activate_page("Chats", TAB_CONTACTS)
		scope = page or root
		if scope is not None and _focus_message_field(scope):
			_announce_found("Message field")
			return
		_announce_missing("Message field")

	@scriptHandler.script(
		gesture="kb:control+2",
		description="Focus the chat history",
		category="Signal Desktop Enhancements",
	)
	def script_focusChatHistory(self, gesture):
		root, page = _activate_page("Chats", TAB_CONTACTS)
		scope = page or root
		if scope is not None and _focus_chat_history(scope):
			_announce_found("Chat history")
			return
		_announce_missing("Chat history")

	@scriptHandler.script(
		gesture="kb:NVDA+shift+d",
		description="Dump Signal accessibility tree to the NVDA log",
		category="Signal Desktop Enhancements",
	)
	def script_dumpTree(self, gesture):
		root = _signal_root()
		if root is None:
			_announce_missing("Signal accessibility tree")
			return
		log.info(_dump_tree(root))
		ui.message("Signal tree written to the NVDA log")
