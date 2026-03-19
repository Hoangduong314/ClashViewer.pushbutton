# -*- coding: utf-8 -*-
import clr
import os
import io
import re
import json
import ntpath
import posixpath
import base64
import System
import traceback

from pyrevit import forms, script, revit, DB
from System.Windows import Window, WindowStartupLocation
from System.Windows.Media import SolidColorBrush, Color

# Load WebView2
clr.AddReference("Microsoft.Web.WebView2.Wpf")
clr.AddReference("Microsoft.Web.WebView2.Core")
from Microsoft.Web.WebView2.Wpf import WebView2

try:
    from urllib import unquote
except ImportError:
    from urllib.parse import unquote

try:
    text_type = unicode
    binary_types = (bytearray,)
except NameError:
    text_type = str
    binary_types = (bytes, bytearray)


class ClashItem(object):
    def __init__(
        self,
        row_key,
        row_html_index,
        name,
        status,
        date,
        item1_id,
        item1_file,
        item2_id,
        item2_file,
        img_rel_path,
    ):
        self.RowKey = row_key
        self.RowHtmlIndex = row_html_index
        self.Name = name
        self.Status = status
        self.Date = date
        self.ImageRelPath = img_rel_path
        self.Item1_ID = item1_id
        self.Item1_File = item1_file
        self.Item2_ID = item2_id
        self.Item2_File = item2_file
        self.DisplayId = item1_id

    def set_display_context(self, selected_nwc=None):
        if selected_nwc:
            if self.Item1_File == selected_nwc:
                self.DisplayId = self.Item1_ID
            elif self.Item2_File == selected_nwc:
                self.DisplayId = self.Item2_ID
            else:
                self.DisplayId = self.Item1_ID
        else:
            self.DisplayId = self.Item1_ID

    def to_view_model(self, html_folder):
        full_img_path = resolve_image_reference(html_folder, self.ImageRelPath)
        return {
            "rowKey": to_text(self.RowKey),
            "name": to_text(self.Name),
            "status": to_text(self.Status),
            "displayId": to_text(self.DisplayId),
            "date": to_text(self.Date),
            "imagePath": to_text(full_img_path),
        }


class ModernWebView(Window):
    def __init__(self):
        Window.__init__(self)

        self.Title = "Modern Clash Viewer"
        self.Width = 1000
        self.Height = 700
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.Background = SolidColorBrush(Color.FromRgb(30, 30, 30))

        self.webView = WebView2()
        self.Content = self.webView

        self.clash_data = []
        self.clash_lookup = {}
        self.original_statuses = {}
        self.html_folder = ""
        self.current_file_path = ""
        self.file_options = set()
        self.selected_nwc_file = None
        self.selected_row_key = None
        self.selected_element_id = None
        self.status_column_index = 1

        temp_cache_folder = os.path.join(os.environ.get("TEMP", os.getcwd()), "pyRevit_WebView2_Cache")
        System.Environment.SetEnvironmentVariable("WEBVIEW2_USER_DATA_FOLDER", temp_cache_folder)

        self.webView.CoreWebView2InitializationCompleted += self.on_webview_ready
        self.webView.EnsureCoreWebView2Async(None)

    def on_webview_ready(self, sender, args):
        if args.IsSuccess:
            html_path = script.get_bundle_file("ui.html")
            file_uri = "file:///" + html_path.replace("\\", "/")
            self.webView.CoreWebView2.Navigate(file_uri)
            self.webView.WebMessageReceived += self.on_message_received
        else:
            forms.alert("Error initializing browser: " + str(args.InitializationException))

    def on_message_received(self, sender, args):
        try:
            payload = self.get_message_payload(args)
            action = payload.get("action")

            if action == "select_file":
                self.SelectFile_Click()
            elif action == "ui_ready":
                self.RestoreConfig()
            elif action == "row_selected":
                self.handle_row_selected(payload)
            elif action == "load_image":
                self.LoadImageBase64(payload.get("image_path", ""))
            elif action == "update_status":
                self.handle_status_update(payload)
            elif action == "save_report":
                self.UpdateReport_Click()
            elif action == "show_in_view":
                self.ShowInView_Click()
        except Exception:
            forms.alert(traceback.format_exc(), title="Data Processing Error")

    def get_message_payload(self, args):
        raw_json = getattr(args, "WebMessageAsJson", None)
        if raw_json:
            payload = json.loads(raw_json)
            if isinstance(payload, dict):
                return payload

        raw_value = args.TryGetWebMessageAsString()
        if not raw_value:
            return {}

        try:
            payload = json.loads(raw_value)
            if isinstance(payload, dict):
                return payload
        except Exception:
            pass

        parts = raw_value.split("|", 1)
        return {
            "action": parts[0],
            "value": parts[1] if len(parts) > 1 else "",
        }

    def handle_row_selected(self, payload):
        self.selected_row_key = to_text(payload.get("row_key", ""))
        self.selected_element_id = to_text(payload.get("selected_id", ""))

        cfg = script.get_config()
        cfg.last_clash_row_key = self.selected_row_key
        cfg.last_clash_id = self.selected_element_id
        script.save_config()

    def handle_status_update(self, payload):
        row_key = to_text(payload.get("row_key", ""))
        new_status = to_text(payload.get("status", ""))
        item = self.clash_lookup.get(row_key)
        if item:
            item.Status = new_status

    def RestoreConfig(self):
        cfg = script.get_config()
        last_file = getattr(cfg, "last_clash_file", "")
        last_nwc = getattr(cfg, "last_nwc_file", "")
        last_row_key = getattr(cfg, "last_clash_row_key", "")

        if last_file and os.path.exists(last_file):
            self.LoadFile(last_file, auto_nwc=last_nwc, auto_select_row_key=last_row_key)

    def call_js(self, function_name, payload=None):
        if payload is None:
            script_text = "{0}();".format(function_name)
        else:
            script_text = "{0}({1});".format(function_name, json.dumps(payload, ensure_ascii=False))
        self.webView.ExecuteScriptAsync(script_text)

    def SelectFile_Click(self):
        file_path = forms.pick_file(file_ext="html")
        if not file_path:
            return
        self.LoadFile(file_path)

    def LoadFile(self, file_path, auto_nwc=None, auto_select_row_key=None):
        self.current_file_path = file_path
        self.html_folder = os.path.dirname(file_path)
        self.selected_row_key = None
        self.selected_element_id = None

        cfg = script.get_config()
        cfg.last_clash_file = file_path
        script.save_config()

        self.call_js("updateFilePath", to_text(file_path))
        self.parse_html_data(file_path)

        if not self.clash_data:
            forms.alert("Could not find clash data in this report.")
            return

        display_list = self.get_display_items(auto_nwc)
        if display_list is None:
            return

        self.call_js("loadClashData", [item.to_view_model(self.html_folder) for item in display_list])

        if auto_select_row_key and auto_select_row_key in self.clash_lookup:
            self.call_js("selectClashByRowKey", to_text(auto_select_row_key))

    def get_display_items(self, auto_nwc=None):
        cfg = script.get_config()
        selected_nwc = None

        if self.file_options:
            sorted_files = sorted(self.file_options)
            if auto_nwc and auto_nwc in self.file_options:
                selected_nwc = auto_nwc
            else:
                selected_nwc = forms.SelectFromList.show(
                    sorted_files,
                    title="Select NWC File to Filter",
                    multiselect=False,
                )
                if not selected_nwc:
                    return None

            self.selected_nwc_file = selected_nwc
            cfg.last_nwc_file = selected_nwc
            script.save_config()
        else:
            self.selected_nwc_file = None
            if hasattr(cfg, "last_nwc_file"):
                cfg.last_nwc_file = ""
                script.save_config()

        display_list = []
        for item in self.clash_data:
            item.set_display_context(self.selected_nwc_file)
            if not self.selected_nwc_file:
                display_list.append(item)
            elif item.Item1_File == self.selected_nwc_file or item.Item2_File == self.selected_nwc_file:
                display_list.append(item)
        return display_list

    def LoadImageBase64(self, img_path):
        try:
            img_path = to_text(img_path).strip()
            if not img_path:
                self.call_js("showBase64Image", "")
                return

            if img_path.startswith("data:"):
                self.call_js("showBase64Image", img_path)
                return

            resolved_img_path = resolve_image_reference(self.html_folder, img_path)
            if not resolved_img_path or not os.path.exists(resolved_img_path):
                self.call_js("showBase64Image", "")
                return

            with io.open(resolved_img_path, "rb") as image_file:
                encoded = base64.b64encode(image_file.read())
                if not isinstance(encoded, text_type):
                    encoded = encoded.decode("utf-8")

            ext = resolved_img_path.lower().rsplit(".", 1)[-1] if "." in resolved_img_path else ""
            mime = "image/jpeg"
            if ext == "png":
                mime = "image/png"
            elif ext == "gif":
                mime = "image/gif"
            elif ext == "bmp":
                mime = "image/bmp"

            self.call_js("showBase64Image", "data:{0};base64,{1}".format(mime, encoded))
        except Exception:
            self.call_js("showBase64Image", "")

    def parse_html_data(self, file_path):
        with io.open(file_path, "r", encoding="utf-8") as report_file:
            content = report_file.read()

        tr_matches = list(re.finditer(r"<tr.*?>.*?</tr>", content, re.DOTALL | re.IGNORECASE))
        th_pattern = re.compile(r"<t[dh].*?>(.*?)</t[dh]>", re.DOTALL | re.IGNORECASE)
        td_pattern = re.compile(r"<td(?:[^>]*/>|[^>]*>(.*?)</td\s*>)", re.DOTALL | re.IGNORECASE)

        headers_raw = []
        for tr_match in tr_matches:
            row_html = tr_match.group(0)
            cells = th_pattern.findall(row_html)
            text_content = " ".join(cells).lower()
            if "status" in text_content and "name" in text_content and "path" in text_content:
                headers_raw = cells
                break

        header_indices = {}
        for index, header in enumerate(headers_raw):
            clean_header = clean_html_text(header).replace("&nbsp;", " ").strip().lower()
            clean_header = clean_header.replace("\r", "").replace("\n", "")
            if not clean_header:
                continue
            header_indices.setdefault(clean_header, []).append(index)

        def get_sorted_indices(keywords, exclude=None):
            exclude = exclude or []
            for keyword in keywords:
                indices = []
                for header_key, idx_list in header_indices.items():
                    if keyword == header_key or keyword in header_key:
                        if not any(ex in header_key for ex in exclude):
                            indices.extend(idx_list)
                if indices:
                    return sorted(set(indices))
            return []

        name_indices = get_sorted_indices(["clash name", "name"], exclude=["item", "group", "file", "path", "document", "system"])
        status_indices = get_sorted_indices(["status"])
        date_indices = get_sorted_indices(["date found", "date"])
        id_indices = get_sorted_indices(["element id", "item id", "guid", "id"], exclude=["grid", "image", "name"])
        path_indices = get_sorted_indices(["path", "file", "document", "layer"], exclude=["grid", "image", "name", "id"])

        name_idx = name_indices[0] if name_indices else 1
        status_idx = status_indices[0] if status_indices else 2
        date_idx = date_indices[0] if date_indices else 4
        id1_idx = id_indices[0] if len(id_indices) > 0 else 5
        id2_idx = id_indices[1] if len(id_indices) > 1 else 10
        path1_idx = path_indices[0] if len(path_indices) > 0 else 7
        path2_idx = path_indices[1] if len(path_indices) > 1 else 12

        self.status_column_index = status_idx
        self.clash_data = []
        self.clash_lookup = {}
        self.file_options = set()

        max_required_idx = max([name_idx, status_idx, date_idx, id1_idx, path1_idx, id2_idx, path2_idx])

        for tr_index, tr_match in enumerate(tr_matches):
            row_html = tr_match.group(0)
            tds = td_pattern.findall(row_html)
            tds = ["" if cell is None else cell for cell in tds]
            if len(tds) <= max_required_idx:
                continue

            img_match = re.search(r"src=[\"\'](.*?)[\"\']", row_html, re.IGNORECASE)
            img_path = img_match.group(1) if img_match else ""

            item = ClashItem(
                row_key="row-{0}".format(len(self.clash_data)),
                row_html_index=tr_index,
                name=clean_html_text(tds[name_idx]),
                status=clean_html_text(tds[status_idx]),
                date=clean_html_text(tds[date_idx]),
                item1_id=clean_html_text(tds[id1_idx]),
                item1_file=extract_filename(clean_html_text(tds[path1_idx])),
                item2_id=clean_html_text(tds[id2_idx]),
                item2_file=extract_filename(clean_html_text(tds[path2_idx])),
                img_rel_path=img_path,
            )
            item.set_display_context()
            self.clash_data.append(item)
            self.clash_lookup[item.RowKey] = item

            if item.Item1_File:
                self.file_options.add(item.Item1_File)
            if item.Item2_File:
                self.file_options.add(item.Item2_File)

        self.original_statuses = dict((item.RowKey, item.Status) for item in self.clash_data)

    def ShowInView_Click(self):
        if not self.selected_element_id:
            forms.alert("Please select a clash from the list first.")
            return

        doc = revit.doc
        uidoc = revit.uidoc
        element = None

        try:
            element = doc.GetElement(DB.ElementId(int(self.selected_element_id)))
        except Exception:
            try:
                element = doc.GetElement(self.selected_element_id)
            except Exception:
                element = None

        if not element:
            forms.alert("Element not found in current model.\nID: {0}".format(self.selected_element_id))
            return

        target_view = None
        if isinstance(doc.ActiveView, DB.View3D) and not doc.ActiveView.IsTemplate:
            target_view = doc.ActiveView
        else:
            views = list(DB.FilteredElementCollector(doc).OfClass(DB.View3D))
            for view in views:
                if view.Name == "{3D}" and not view.IsTemplate:
                    target_view = view
                    break
            if not target_view:
                for view in views:
                    if not view.IsTemplate:
                        target_view = view
                        break

        if not target_view:
            forms.alert("No suitable 3D View found.")
            return

        transaction = None
        try:
            transaction = DB.Transaction(doc, "Clash Viewer Section Box")
            transaction.Start()

            bbox = element.get_BoundingBox(None)
            if not bbox:
                forms.alert("Element has no geometry.")
                transaction.RollBack()
                return

            offset = 1.6
            section_box = DB.BoundingBoxXYZ()
            section_box.Min = DB.XYZ(bbox.Min.X - offset, bbox.Min.Y - offset, bbox.Min.Z - offset)
            section_box.Max = DB.XYZ(bbox.Max.X + offset, bbox.Max.Y + offset, bbox.Max.Z + offset)
            target_view.SetSectionBox(section_box)
            target_view.IsSectionBoxActive = True
            transaction.Commit()

            if doc.ActiveView.Id != target_view.Id:
                uidoc.RequestViewChange(target_view)

            selection_ids = System.Collections.Generic.List[DB.ElementId]()
            selection_ids.Add(element.Id)
            uidoc.Selection.SetElementIds(selection_ids)
            uidoc.ShowElements(element.Id)
        except Exception as exc:
            if transaction and transaction.GetStatus() == DB.TransactionStatus.Started:
                transaction.RollBack()
            forms.alert("Error processing view: " + str(exc))

    def UpdateReport_Click(self):
        if not self.current_file_path or not self.clash_data:
            return

        changed_items = [item for item in self.clash_data if self.original_statuses.get(item.RowKey) != item.Status]
        if not changed_items:
            forms.alert("No status changes to save.", title="No Changes")
            return

        try:
            with io.open(self.current_file_path, "r", encoding="utf-8") as report_file:
                content = report_file.read()

            tr_matches = list(re.finditer(r"<tr.*?>.*?</tr>", content, re.DOTALL | re.IGNORECASE))
            row_updates = {}
            for item in changed_items:
                if item.RowHtmlIndex >= len(tr_matches):
                    continue
                original_row = tr_matches[item.RowHtmlIndex].group(0)
                row_updates[item.RowHtmlIndex] = self.replace_table_cell(original_row, self.status_column_index, item.Status)

            if not row_updates:
                forms.alert("Could not map clash rows back to the report.", title="Save Failed")
                return

            rebuilt_parts = []
            last_end = 0
            for row_index, tr_match in enumerate(tr_matches):
                start, end = tr_match.span()
                rebuilt_parts.append(content[last_end:start])
                rebuilt_parts.append(row_updates.get(row_index, tr_match.group(0)))
                last_end = end
            rebuilt_parts.append(content[last_end:])
            updated_content = "".join(rebuilt_parts)

            with io.open(self.current_file_path, "w", encoding="utf-8") as report_file:
                report_file.write(updated_content)

            self.original_statuses = dict((item.RowKey, item.Status) for item in self.clash_data)
            forms.alert("Report saved successfully!", title="Saved")
        except Exception as exc:
            forms.alert("Failed to save: {0}".format(str(exc)))

    def replace_table_cell(self, row_html, cell_index, new_value):
        td_matches = list(re.finditer(r"<td\b([^>]*)>(.*?)</td\s*>", row_html, re.DOTALL | re.IGNORECASE))
        if cell_index >= len(td_matches):
            return row_html

        target_match = td_matches[cell_index]
        attrs = target_match.group(1) or ""
        replacement = "<td{0}>{1}</td>".format(attrs, escape_html(to_text(new_value)))
        start, end = target_match.span()
        return row_html[:start] + replacement + row_html[end:]


def to_text(value):
    if value is None:
        return u""
    if isinstance(value, text_type):
        return value
    if isinstance(value, System.String):
        return text_type(value)
    if isinstance(value, binary_types):
        try:
            return bytes(value).decode("utf-8", "ignore")
        except Exception:
            return text_type(value)
    try:
        return value.decode("utf-8", "ignore")
    except Exception:
        return text_type(value)


def normalize_path(path_value):
    return to_text(path_value).replace("/", os.sep).replace("\\", os.sep).replace(u"\xa0", u" ").strip()


def resolve_image_reference(base_folder, image_ref):
    image_ref = to_text(image_ref).strip().strip("\"'")
    if not image_ref:
        return u""

    image_ref = image_ref.replace("&amp;", "&")
    image_ref = unquote(image_ref)

    if image_ref.startswith("data:"):
        return image_ref

    if image_ref.lower().startswith("file:///"):
        image_ref = image_ref[8:]
    elif image_ref.lower().startswith("file://"):
        image_ref = image_ref[7:]

    if re.match(r'^[a-zA-Z]:[\\/]', image_ref) or image_ref.startswith('\\'):
        return os.path.normpath(image_ref.replace('/', os.sep))

    normalized_rel = posixpath.normpath(image_ref.replace('\\', '/'))
    normalized_rel = normalized_rel.lstrip('/')
    if normalized_rel and normalized_rel != '.':
        return os.path.normpath(os.path.join(base_folder, normalized_rel.replace('/', os.sep)))

    return os.path.normpath(os.path.join(base_folder, ntpath.basename(image_ref)))


def clean_html_text(text):
    clean = re.sub(r"<.*?>", "", to_text(text))
    clean = clean.replace("&nbsp;", " ")
    clean = clean.replace("&gt;", ">")
    clean = clean.replace("&lt;", "<")
    clean = clean.replace("&amp;", "&")
    clean = clean.strip()
    clean = clean.replace("Element ID", "").replace("GUID", "").replace(":", "").strip()
    return clean


def extract_filename(path_str):
    normalized = to_text(path_str)
    match = re.search(r"File\s*>\s*File\s*>\s*(.*?)\.nwc", normalized, re.IGNORECASE)
    if match:
        return match.group(1).strip() + ".nwc"
    return "Unknown File"


def escape_html(value):
    text = to_text(value)
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


if __name__ == "__main__":
    try:
        app = ModernWebView()
        app.ShowDialog()
    except Exception:
        forms.alert(traceback.format_exc(), title="Script Startup Error")
