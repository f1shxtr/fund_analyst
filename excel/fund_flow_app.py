import csv
import datetime as dt
import math
import os
import re
import sys
import tkinter as tk
import zipfile
from collections import defaultdict
from dataclasses import dataclass
from tkinter import filedialog, messagebox, ttk
from xml.etree import ElementTree as ET


NS_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
NS_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
NS_PKG_REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"


def xml_path(name):
    normalized = name.lstrip("/")
    if normalized.startswith("xl/"):
        return normalized
    return "xl/" + normalized


def col_to_index(cell_ref):
    match = re.match(r"([A-Z]+)", cell_ref or "")
    if not match:
        return 0
    value = 0
    for char in match.group(1):
        value = value * 26 + (ord(char) - ord("A") + 1)
    return value - 1


def norm_person(value):
    text = str(value or "").strip()
    text = re.sub(r"\s+", "", text)
    text = text.replace("（", "(").replace("）", ")")
    return text.upper()


def parse_amount(value):
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None
    text = text.replace(",", "").replace("￥", "").replace("¥", "")
    text = re.sub(r"[^\d.\-]", "", text)
    if text in {"", "-", ".", "-."}:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def excel_serial_to_date(value):
    try:
        serial = float(value)
    except (TypeError, ValueError):
        return None
    if not 1 <= serial <= 90000:
        return None
    base = dt.datetime(1899, 12, 30)
    return (base + dt.timedelta(days=serial)).date()


def parse_date(value):
    if isinstance(value, dt.date):
        return value
    if value is None:
        return None

    text = str(value).strip()
    if not text:
        return None

    serial_date = excel_serial_to_date(text)
    if serial_date and re.fullmatch(r"\d+(\.\d+)?", text):
        return serial_date

    text = text.replace("年", "-").replace("月", "-").replace("日", "")
    text = text.replace("/", "-").replace(".", "-")
    text = re.sub(r"\s+\d{1,2}:\d{2}(:\d{2})?.*$", "", text)
    text = text.split("T")[0]

    for fmt in ("%Y-%m-%d", "%Y-%m-%d %H:%M:%S", "%Y%m%d"):
        try:
            return dt.datetime.strptime(text, fmt).date()
        except ValueError:
            pass

    match = re.search(r"(\d{4})-(\d{1,2})-(\d{1,2})", text)
    if match:
        try:
            return dt.date(int(match.group(1)), int(match.group(2)), int(match.group(3)))
        except ValueError:
            return None
    return None


def fmt_amount(value):
    if value is None:
        return ""
    return f"{value:,.2f}"


def fmt_date(value):
    return value.isoformat() if isinstance(value, dt.date) else ""


class SimpleXlsx:
    def __init__(self, path):
        self.path = path
        self.shared_strings = []
        self.sheets = []
        self._load()

    def _load(self):
        with zipfile.ZipFile(self.path) as zf:
            self.shared_strings = self._read_shared_strings(zf)
            rels = self._read_workbook_rels(zf)
            root = ET.fromstring(zf.read("xl/workbook.xml"))
            sheets_node = root.find(NS_MAIN + "sheets")
            if sheets_node is None:
                return
            for sheet in sheets_node.findall(NS_MAIN + "sheet"):
                name = sheet.attrib.get("name", "Sheet")
                rel_id = sheet.attrib.get(NS_REL + "id")
                target = rels.get(rel_id)
                if target:
                    self.sheets.append((name, xml_path(target)))

    def _read_shared_strings(self, zf):
        if "xl/sharedStrings.xml" not in zf.namelist():
            return []
        root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
        values = []
        for item in root.findall(NS_MAIN + "si"):
            parts = []
            for text in item.iter(NS_MAIN + "t"):
                parts.append(text.text or "")
            values.append("".join(parts))
        return values

    def _read_workbook_rels(self, zf):
        rels_path = "xl/_rels/workbook.xml.rels"
        if rels_path not in zf.namelist():
            return {}
        root = ET.fromstring(zf.read(rels_path))
        rels = {}
        for rel in root.findall(NS_PKG_REL + "Relationship"):
            rels[rel.attrib.get("Id")] = rel.attrib.get("Target", "")
        return rels

    def read_sheet(self, sheet_name):
        path = None
        for name, sheet_path in self.sheets:
            if name == sheet_name:
                path = sheet_path
                break
        if not path:
            raise ValueError("找不到工作表")

        with zipfile.ZipFile(self.path) as zf:
            root = ET.fromstring(zf.read(path))
        data = root.find(NS_MAIN + "sheetData")
        if data is None:
            return []

        rows = []
        for row in data.findall(NS_MAIN + "row"):
            values = []
            for cell in row.findall(NS_MAIN + "c"):
                idx = col_to_index(cell.attrib.get("r", ""))
                while len(values) <= idx:
                    values.append("")
                values[idx] = self._cell_value(cell)
            rows.append(values)

        width = max((len(row) for row in rows), default=0)
        return [row + [""] * (width - len(row)) for row in rows]

    def _cell_value(self, cell):
        cell_type = cell.attrib.get("t")
        if cell_type == "inlineStr":
            parts = []
            inline = cell.find(NS_MAIN + "is")
            if inline is not None:
                for text in inline.iter(NS_MAIN + "t"):
                    parts.append(text.text or "")
            return "".join(parts)

        value_node = cell.find(NS_MAIN + "v")
        if value_node is None:
            return ""
        value = value_node.text or ""

        if cell_type == "s":
            try:
                return self.shared_strings[int(value)]
            except (ValueError, IndexError):
                return value
        if cell_type == "b":
            return "TRUE" if value == "1" else "FALSE"
        return value


def read_csv(path):
    for encoding in ("utf-8-sig", "gb18030", "utf-8"):
        try:
            with open(path, "r", encoding=encoding, newline="") as handle:
                return list(csv.reader(handle))
        except UnicodeDecodeError:
            continue
    with open(path, "r", encoding="utf-8", errors="replace", newline="") as handle:
        return list(csv.reader(handle))


def first_header_row(rows):
    for idx, row in enumerate(rows[:20]):
        non_empty = [str(cell).strip() for cell in row if str(cell).strip()]
        if len(non_empty) >= 3:
            return idx
    return 0


def clean_headers(row):
    headers = []
    seen = defaultdict(int)
    for idx, cell in enumerate(row):
        name = str(cell).strip() or f"列{idx + 1}"
        seen[name] += 1
        if seen[name] > 1:
            name = f"{name}_{seen[name]}"
        headers.append(name)
    return headers


def guess_column(headers, keywords):
    scored = []
    for idx, header in enumerate(headers):
        compact = norm_person(header)
        score = 0
        for keyword in keywords:
            if norm_person(keyword) in compact:
                score += len(keyword)
        scored.append((score, -idx, header))
    scored.sort(reverse=True)
    return scored[0][2] if scored and scored[0][0] > 0 else ""


@dataclass
class Transaction:
    tx_id: int
    row_number: int
    date: dt.date
    sender: str
    receiver: str
    amount: float | None
    raw: dict


def build_transactions(rows, headers, mapping, header_idx):
    transactions = []
    errors = []
    col_index = {name: idx for idx, name in enumerate(headers)}

    for offset, row in enumerate(rows[header_idx + 1 :], start=header_idx + 2):
        row_dict = {header: row[idx] if idx < len(row) else "" for header, idx in col_index.items()}
        date = parse_date(row_dict.get(mapping["date"]))
        sender = str(row_dict.get(mapping["sender"], "")).strip()
        receiver = str(row_dict.get(mapping["receiver"], "")).strip()
        amount = parse_amount(row_dict.get(mapping["amount"])) if mapping.get("amount") else None

        if not any(str(cell).strip() for cell in row):
            continue
        if not date or not sender or not receiver:
            errors.append(f"第 {offset} 行缺少日期/汇出方/汇入方，已跳过")
            continue
        transactions.append(
            Transaction(
                tx_id=len(transactions),
                row_number=offset,
                date=date,
                sender=sender,
                receiver=receiver,
                amount=amount,
                raw=row_dict,
            )
        )
    return transactions, errors


def amount_matches(prev_amount, next_amount, tolerance_percent):
    if tolerance_percent is None or prev_amount is None or next_amount is None:
        return True
    if prev_amount == 0:
        return next_amount == 0
    return abs(next_amount - prev_amount) / abs(prev_amount) <= tolerance_percent / 100


def analyze_flows(transactions, window_days=2, max_depth=5, tolerance_percent=None, max_links=20000, max_chains=5000):
    by_sender = defaultdict(list)
    for tx in transactions:
        by_sender[norm_person(tx.sender)].append(tx)
    for txs in by_sender.values():
        txs.sort(key=lambda item: (item.date, item.row_number))

    links = []
    adjacency = defaultdict(list)
    for tx in sorted(transactions, key=lambda item: (item.date, item.row_number)):
        candidates = by_sender.get(norm_person(tx.receiver), [])
        for nxt in candidates:
            if nxt.tx_id == tx.tx_id:
                continue
            delta = (nxt.date - tx.date).days
            if delta < 0:
                continue
            if delta > window_days:
                if nxt.date > tx.date:
                    break
                continue
            if not amount_matches(tx.amount, nxt.amount, tolerance_percent):
                continue
            link = {
                "prev": tx,
                "next": nxt,
                "days": delta,
                "ratio": "" if tx.amount in (None, 0) or nxt.amount is None else nxt.amount / tx.amount,
            }
            links.append(link)
            adjacency[tx.tx_id].append(link)
            if len(links) >= max_links:
                break
        if len(links) >= max_links:
            break

    chains = []

    def dfs(path, seen):
        last = path[-1]
        next_links = [link for link in adjacency.get(last.tx_id, []) if link["next"].tx_id not in seen]
        if not next_links or len(path) >= max_depth:
            if len(path) >= 2:
                chains.append(path[:])
            return
        for link in next_links:
            if len(chains) >= max_chains:
                return
            dfs(path + [link["next"]], seen | {link["next"].tx_id})

    starts = sorted(transactions, key=lambda item: (item.date, item.row_number))
    for tx in starts:
        if len(chains) >= max_chains:
            break
        if adjacency.get(tx.tx_id):
            dfs([tx], {tx.tx_id})

    return links, chains


def write_csv(path, rows, headers):
    with open(path, "w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=headers)
        writer.writeheader()
        writer.writerows(rows)


def export_results(base_dir, links, chains):
    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    link_path = os.path.join(base_dir, f"资金流向相邻关系_{stamp}.csv")
    chain_path = os.path.join(base_dir, f"资金流向链条_{stamp}.csv")

    link_rows = []
    for idx, link in enumerate(links, start=1):
        prev = link["prev"]
        nxt = link["next"]
        ratio = link["ratio"]
        link_rows.append(
            {
                "序号": idx,
                "前笔原表行号": prev.row_number,
                "后笔原表行号": nxt.row_number,
                "前笔日期": fmt_date(prev.date),
                "中间人": prev.receiver,
                "前笔汇出方": prev.sender,
                "后笔汇入方": nxt.receiver,
                "前笔金额": fmt_amount(prev.amount),
                "后笔金额": fmt_amount(nxt.amount),
                "间隔天数": link["days"],
                "后笔/前笔金额比例": "" if ratio == "" else f"{ratio:.4f}",
            }
        )

    chain_rows = []
    for chain_id, chain in enumerate(chains, start=1):
        for step, tx in enumerate(chain, start=1):
            days_from_prev = ""
            if step > 1:
                days_from_prev = (tx.date - chain[step - 2].date).days
            chain_rows.append(
                {
                    "链条编号": chain_id,
                    "步骤": step,
                    "原表行号": tx.row_number,
                    "日期": fmt_date(tx.date),
                    "汇出方": tx.sender,
                    "汇入方": tx.receiver,
                    "金额": fmt_amount(tx.amount),
                    "与上一步间隔天数": days_from_prev,
                    "链条路径": " -> ".join([chain[0].sender] + [item.receiver for item in chain]),
                }
            )

    write_csv(
        link_path,
        link_rows,
        ["序号", "前笔原表行号", "后笔原表行号", "前笔日期", "中间人", "前笔汇出方", "后笔汇入方", "前笔金额", "后笔金额", "间隔天数", "后笔/前笔金额比例"],
    )
    write_csv(
        chain_path,
        chain_rows,
        ["链条编号", "步骤", "原表行号", "日期", "汇出方", "汇入方", "金额", "与上一步间隔天数", "链条路径"],
    )
    return link_path, chain_path


class FundFlowApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("本地资金流向链条分析")
        self.geometry("1060x720")
        self.minsize(960, 620)

        self.file_path = ""
        self.workbook = None
        self.rows = []
        self.header_idx = 0
        self.headers = []

        self.file_var = tk.StringVar()
        self.sheet_var = tk.StringVar()
        self.date_col = tk.StringVar()
        self.sender_col = tk.StringVar()
        self.receiver_col = tk.StringVar()
        self.amount_col = tk.StringVar()
        self.window_var = tk.StringVar(value="2")
        self.depth_var = tk.StringVar(value="5")
        self.tolerance_var = tk.StringVar(value="")
        self.status_var = tk.StringVar(value="请选择 .xlsx 或 .csv 文件")

        self._build_ui()

    def _build_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(3, weight=1)

        top = ttk.Frame(self, padding=(12, 10, 12, 6))
        top.grid(row=0, column=0, sticky="ew")
        top.columnconfigure(1, weight=1)

        ttk.Button(top, text="选择文件", command=self.choose_file).grid(row=0, column=0, sticky="w")
        ttk.Entry(top, textvariable=self.file_var, state="readonly").grid(row=0, column=1, sticky="ew", padx=8)
        ttk.Label(top, text="工作表").grid(row=0, column=2, padx=(8, 4))
        self.sheet_combo = ttk.Combobox(top, textvariable=self.sheet_var, state="readonly", width=18)
        self.sheet_combo.grid(row=0, column=3, sticky="e")
        self.sheet_combo.bind("<<ComboboxSelected>>", lambda _event: self.load_selected_sheet())

        mapping = ttk.LabelFrame(self, text="字段对应", padding=12)
        mapping.grid(row=1, column=0, sticky="ew", padx=12, pady=6)
        for idx in range(8):
            mapping.columnconfigure(idx, weight=1 if idx % 2 else 0)

        self.date_combo = self._combo(mapping, "日期", self.date_col, 0)
        self.sender_combo = self._combo(mapping, "汇出方/付款人", self.sender_col, 2)
        self.receiver_combo = self._combo(mapping, "汇入方/收款人", self.receiver_col, 4)
        self.amount_combo = self._combo(mapping, "金额", self.amount_col, 6)

        options = ttk.LabelFrame(self, text="链条规则", padding=12)
        options.grid(row=2, column=0, sticky="ew", padx=12, pady=6)
        ttk.Label(options, text="相邻天数").grid(row=0, column=0, sticky="w")
        ttk.Spinbox(options, from_=0, to=30, width=8, textvariable=self.window_var).grid(row=0, column=1, sticky="w", padx=(6, 18))
        ttk.Label(options, text="最大链长").grid(row=0, column=2, sticky="w")
        ttk.Spinbox(options, from_=2, to=20, width=8, textvariable=self.depth_var).grid(row=0, column=3, sticky="w", padx=(6, 18))
        ttk.Label(options, text="金额近似容差 %").grid(row=0, column=4, sticky="w")
        ttk.Entry(options, width=10, textvariable=self.tolerance_var).grid(row=0, column=5, sticky="w", padx=(6, 18))
        ttk.Button(options, text="分析并导出 CSV", command=self.run_analysis).grid(row=0, column=6, sticky="e")
        options.columnconfigure(7, weight=1)

        results = ttk.Frame(self, padding=(12, 6, 12, 6))
        results.grid(row=3, column=0, sticky="nsew")
        results.rowconfigure(0, weight=1)
        results.columnconfigure(0, weight=1)

        columns = ("chain", "path", "dates", "amounts", "rows")
        self.tree = ttk.Treeview(results, columns=columns, show="headings")
        self.tree.heading("chain", text="链条编号")
        self.tree.heading("path", text="链条路径")
        self.tree.heading("dates", text="日期")
        self.tree.heading("amounts", text="金额")
        self.tree.heading("rows", text="原表行号")
        self.tree.column("chain", width=80, anchor="center")
        self.tree.column("path", width=440)
        self.tree.column("dates", width=180)
        self.tree.column("amounts", width=180)
        self.tree.column("rows", width=120)
        self.tree.grid(row=0, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(results, orient="vertical", command=self.tree.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=scrollbar.set)

        status = ttk.Frame(self, padding=(12, 4, 12, 10))
        status.grid(row=4, column=0, sticky="ew")
        status.columnconfigure(0, weight=1)
        ttk.Label(status, textvariable=self.status_var).grid(row=0, column=0, sticky="w")

    def _combo(self, parent, label, var, col):
        ttk.Label(parent, text=label).grid(row=0, column=col, sticky="w")
        combo = ttk.Combobox(parent, textvariable=var, state="readonly", width=20)
        combo.grid(row=0, column=col + 1, sticky="ew", padx=(6, 12))
        return combo

    def choose_file(self):
        path = filedialog.askopenfilename(
            title="选择交易流水",
            filetypes=[("Excel 工作簿", "*.xlsx"), ("CSV 文件", "*.csv"), ("所有文件", "*.*")],
        )
        if not path:
            return
        self.file_path = path
        self.file_var.set(path)
        try:
            if path.lower().endswith(".xlsx"):
                self.workbook = SimpleXlsx(path)
                sheet_names = [name for name, _path in self.workbook.sheets]
                if not sheet_names:
                    raise ValueError("未读取到工作表")
                self.sheet_combo.configure(values=sheet_names, state="readonly")
                self.sheet_var.set(sheet_names[0])
                self.load_selected_sheet()
            else:
                self.workbook = None
                self.sheet_combo.configure(values=["CSV"], state="disabled")
                self.sheet_var.set("CSV")
                self.rows = read_csv(path)
                self.prepare_headers()
        except Exception as exc:
            messagebox.showerror("读取失败", str(exc))
            self.status_var.set("读取失败")

    def load_selected_sheet(self):
        if not self.workbook:
            return
        self.rows = self.workbook.read_sheet(self.sheet_var.get())
        self.prepare_headers()

    def prepare_headers(self):
        if not self.rows:
            raise ValueError("文件没有可读取的数据")
        self.header_idx = first_header_row(self.rows)
        self.headers = clean_headers(self.rows[self.header_idx])
        values = self.headers
        for combo in (self.date_combo, self.sender_combo, self.receiver_combo, self.amount_combo):
            combo.configure(values=values)

        self.date_col.set(guess_column(values, ["交易日期", "发生日期", "入账日期", "日期", "时间", "DATE"]))
        self.sender_col.set(guess_column(values, ["汇出方", "付款人", "付款方", "转出方", "转出户名", "借方户名", "SENDER", "FROM"]))
        self.receiver_col.set(guess_column(values, ["汇入方", "收款人", "收款方", "转入方", "转入户名", "贷方户名", "RECEIVER", "TO"]))
        self.amount_col.set(guess_column(values, ["交易金额", "发生额", "金额", "转账金额", "AMOUNT"]))

        self.status_var.set(f"已读取 {len(self.rows) - self.header_idx - 1} 行数据；请确认字段对应后分析")

    def run_analysis(self):
        try:
            if not self.rows:
                raise ValueError("请先选择文件")
            mapping = {
                "date": self.date_col.get(),
                "sender": self.sender_col.get(),
                "receiver": self.receiver_col.get(),
                "amount": self.amount_col.get(),
            }
            missing = [name for name, value in mapping.items() if name != "amount" and not value]
            if missing:
                raise ValueError("请先选择日期、汇出方、汇入方字段")

            window_days = int(self.window_var.get())
            max_depth = int(self.depth_var.get())
            tolerance = self.parse_tolerance()

            transactions, errors = build_transactions(self.rows, self.headers, mapping, self.header_idx)
            links, chains = analyze_flows(transactions, window_days, max_depth, tolerance)
            self.populate_results(chains)

            base_dir = os.path.dirname(self.file_path) or os.getcwd()
            link_path, chain_path = export_results(base_dir, links, chains)
            skipped = f"；跳过 {len(errors)} 行" if errors else ""
            self.status_var.set(f"完成：{len(transactions)} 笔有效流水，{len(links)} 个相邻关系，{len(chains)} 条链条{skipped}")
            messagebox.showinfo("分析完成", f"已导出：\n{link_path}\n{chain_path}")
        except Exception as exc:
            messagebox.showerror("分析失败", str(exc))
            self.status_var.set("分析失败")

    def parse_tolerance(self):
        text = self.tolerance_var.get().strip()
        if not text:
            return None
        value = float(text)
        if math.isnan(value) or value < 0:
            raise ValueError("金额近似容差必须是非负数字")
        return value

    def populate_results(self, chains):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for idx, chain in enumerate(chains[:300], start=1):
            path = " -> ".join([chain[0].sender] + [tx.receiver for tx in chain])
            dates = " / ".join(fmt_date(tx.date) for tx in chain)
            amounts = " / ".join(fmt_amount(tx.amount) for tx in chain)
            rows = " / ".join(str(tx.row_number) for tx in chain)
            self.tree.insert("", "end", values=(idx, path, dates, amounts, rows))


def main():
    app = FundFlowApp()
    app.mainloop()


if __name__ == "__main__":
    main()
