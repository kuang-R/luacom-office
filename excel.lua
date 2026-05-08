---@class ExcelRange
---Excel 单元格/区域操作对象
local ExcelRange = {}
ExcelRange.__index = ExcelRange

---@class ExcelSheet
---Excel 工作表操作对象
local ExcelSheet = {}
ExcelSheet.__index = ExcelSheet

---@class ExcelWorkbook
---Excel 工作簿操作对象
local ExcelWorkbook = {}
ExcelWorkbook.__index = ExcelWorkbook

---@class ExcelApp
---Excel 应用程序操作对象
local ExcelApp = {}
ExcelApp.__index = ExcelApp

-- ============================================================
-- ExcelRange 方法
-- ============================================================

---@param com_range userdata COM Range 对象
---@return ExcelRange
function ExcelRange:new(com_range)
    local obj = { _com = com_range }
    setmetatable(obj, self)
    return obj
end

---获取或设置单元格/区域的值
---
---设置时 nil 等价于清空单元格。
---
---```lua
---local v = range:value()       -- 读取
---range:value("Hello")          -- 写入
---range:value(nil)              -- 清空
---```
---
---@overload fun():any
---@overload fun(v:any):ExcelRange
function ExcelRange:value(v)
    if v == nil then
        v = ""  -- nil 在 COM 中可能报错，用空串代替清空
    end
    if v ~= nil then
        if type(v) == "table" and v.year then
            v = ExcelApp.dateSerial(v)  -- 日期 table → Excel 序列号
        end
        self._com.Value2 = v
        return self
    else
        return self._com.Value2
    end
end

---获取或设置单元格公式
---@overload fun():string
---@overload fun(formula:string):ExcelRange
function ExcelRange:formula(f)
    if f ~= nil then
        self._com.Formula = f
        return self
    else
        return self._com.Formula
    end
end

---获取行数
---@return integer
function ExcelRange:rows()
    return self._com.Rows.Count
end

---获取列数
---@return integer
function ExcelRange:columns()
    return self._com.Columns.Count
end

---获取区域地址字符串，如 "$A$1:$C$10"
---@return string
function ExcelRange:address()
    return self._com.Address
end

---获取区域左上角单元格的行号 (1-based)
---@return integer
function ExcelRange:row()
    return self._com.Row
end

---获取区域左上角单元格的列号 (1-based)
---@return integer
function ExcelRange:column()
    return self._com.Column
end

---获取指定单元格 (1-based)
---@param r integer 行
---@param c integer 列
---@return ExcelRange
function ExcelRange:cell(r, c)
    return ExcelRange:new(self._com.Cells(r, c))
end

---在当前区域内偏移
---@param row_offset integer
---@param col_offset integer
---@return ExcelRange
function ExcelRange:offset(row_offset, col_offset)
    return ExcelRange:new(self._com.Offset(row_offset, col_offset))
end

---获取整行
---@return ExcelRange
function ExcelRange:entireRow()
    return ExcelRange:new(self._com.EntireRow)
end

---获取整列
---@return ExcelRange
function ExcelRange:entireColumn()
    return ExcelRange:new(self._com.EntireColumn)
end

-- ============================================================
-- 字体 (Font)
-- ============================================================

---设置字体名称
---@param name string
---@return ExcelRange
function ExcelRange:fontName(name)
    self._com.Font.Name = name
    return self
end

---设置字体大小
---@param size number
---@return ExcelRange
function ExcelRange:fontSize(size)
    self._com.Font.Size = size
    return self
end

---设置粗体
---@param enable boolean|nil 不传参则读取
---@return boolean|ExcelRange
function ExcelRange:bold(enable)
    if enable ~= nil then
        self._com.Font.Bold = enable
        return self
    else
        return self._com.Font.Bold
    end
end

---设置斜体
---@param enable boolean|nil
---@return boolean|ExcelRange
function ExcelRange:italic(enable)
    if enable ~= nil then
        self._com.Font.Italic = enable
        return self
    else
        return self._com.Font.Italic
    end
end

---设置下划线
---@param enable boolean|nil
---@return boolean|ExcelRange
function ExcelRange:underline(enable)
    if enable ~= nil then
        self._com.Font.Underline = enable
        return self
    else
        return self._com.Font.Underline
    end
end

---设置字体颜色 (RGB 数值)
---@param rgb integer 例: 0xFF0000 红色
---@return ExcelRange
function ExcelRange:fontColor(rgb)
    self._com.Font.Color = rgb
    return self
end

-- ============================================================
-- Excel COM 常量映射
-- 同时接受字符串名和整数；字符串自动查表转换
-- ============================================================
local XlHAlign = {
    General = 1, Left = -4131, Center = -4108, Right = -4152,
    Fill = 5, Justify = -4130, CenterAcrossSelection = 7, Distributed = -4117,
}
local XlVAlign = {
    Top = -4160, Center = -4108, Bottom = -4107,
    Justify = -4130, Distributed = -4117,
}
local XlLineStyle = {
    Continuous = 1, Dash = -4115, Dot = -4118, Double = -4119,
    None = -4142, SlantDashDot = 13,
}
local XlPattern = {
    Solid = 1, None = -4142, Gray75 = 12, Gray50 = 13, Gray25 = 14,
}

-- ============================================================
-- 对齐 (Alignment)
-- ============================================================

---水平对齐
---
---可选: "General"|"Left"|"Center"|"Right"|"Fill"|"Justify"|"CenterAcrossSelection"|"Distributed" 或整数
---
---```lua
---range:halign("Center")
---range:halign(-4108)  -- 等价
---```
---
---@param align string|integer
---@return ExcelRange
function ExcelRange:halign(align)
    self._com.HorizontalAlignment = XlHAlign[align] or align
    return self
end

---垂直对齐
---
---可选: "Top"|"Center"|"Bottom"|"Justify"|"Distributed" 或整数
---
---@param align string|integer
---@return ExcelRange
function ExcelRange:valign(align)
    self._com.VerticalAlignment = XlVAlign[align] or align
    return self
end

---自动换行
---@param enable boolean
---@return ExcelRange
function ExcelRange:wrapText(enable)
    self._com.WrapText = enable
    return self
end

-- ============================================================
-- 背景/填充 (Interior)
-- ============================================================

---设置背景颜色 (RGB 数值)
---@param rgb integer
---@return ExcelRange
function ExcelRange:bgColor(rgb)
    self._com.Interior.Color = rgb
    return self
end

---设置背景图案
---可选: "Solid"|"None"|"Gray75"|"Gray50"|"Gray25" 或整数
---@param pattern string|integer
---@return ExcelRange
function ExcelRange:bgPattern(pattern)
    self._com.Interior.Pattern = XlPattern[pattern] or pattern
    return self
end

-- ============================================================
-- 边框 (Borders)
-- ============================================================

---设置边框
---
---border_index: 7=左, 8=上, 9=下, 10=右, 11=内部竖线, 12=内部横线
---
---line_style: "Continuous"|"Dash"|"Dot"|"Double"|"None"|"SlantDashDot" 或整数
---
---```lua
---range:border(7, "Continuous")  -- 左边框实线
---range:border(7, 1)             -- 等价
---range:borderAll("Continuous")  -- 全部边框
---```
---
---@param border_index integer 边框索引 (7-12)
---@param line_style string|integer 线型
---@return ExcelRange
function ExcelRange:border(border_index, line_style)
    self._com.Borders(border_index).LineStyle = XlLineStyle[line_style] or line_style
    return self
end

---设置全部边框
---@param line_style string|integer
---@return ExcelRange
function ExcelRange:borderAll(line_style)
    local ls = XlLineStyle[line_style] or line_style
    for i = 7, 12 do
        self._com.Borders(i).LineStyle = ls
    end
    return self
end

---设置外边框
---@param line_style string|integer
---@return ExcelRange
function ExcelRange:borderOutline(line_style)
    local ls = XlLineStyle[line_style] or line_style
    for i = 7, 10 do
        self._com.Borders(i).LineStyle = ls
    end
    return self
end

-- ============================================================
-- 数字格式
-- ============================================================

---设置数字格式
---
---```lua
---range:numberFormat("0.00")        -- 两位小数
---range:numberFormat("yyyy-mm-dd")  -- 日期格式
---range:numberFormat("@")           -- 文本格式
---```
---
---@param fmt string
---@return ExcelRange
function ExcelRange:numberFormat(fmt)
    self._com.NumberFormat = fmt
    return self
end

-- ============================================================
-- 列宽/行高
-- ============================================================

---设置列宽 (仅对整列有效)
---@param w number
---@return ExcelRange
function ExcelRange:columnWidth(w)
    self._com.ColumnWidth = w
    return self
end

---自动调整列宽
---@return ExcelRange
function ExcelRange:autoFit()
    self._com.EntireColumn:AutoFit()
    return self
end

---设置行高
---@param h number
---@return ExcelRange
function ExcelRange:rowHeight(h)
    self._com.RowHeight = h
    return self
end

-- ============================================================
-- 合并
-- ============================================================

---合并单元格
---@return ExcelRange
function ExcelRange:merge()
    self._com:Merge()
    return self
end

---取消合并
---@return ExcelRange
function ExcelRange:unmerge()
    self._com:UnMerge()
    return self
end

-- ============================================================
-- 复制/剪切/粘贴
-- ============================================================

---复制
---@return ExcelRange
function ExcelRange:copy()
    self._com:Copy()
    return self
end

---剪切
---@return ExcelRange
function ExcelRange:cut()
    self._com:Cut()
    return self
end

---粘贴 (粘贴到目标 Range 上)
---@param target ExcelRange
---@return ExcelRange
function ExcelRange:pasteTo(target)
    self._com:Copy(target._com)
    return self
end

-- ============================================================
-- 清除
-- ============================================================

---清除内容
---@return ExcelRange
function ExcelRange:clear()
    self._com:ClearContents()
    return self
end

---清除全部 (含格式)
---@return ExcelRange
function ExcelRange:clearAll()
    self._com:Clear()
    return self
end

-- ============================================================
-- 插入/删除
-- ============================================================

---插入行 (向下移位)
---@return ExcelRange
function ExcelRange:insertRows()
    self._com.EntireRow:Insert()
    return self
end

---删除行
---@return ExcelRange
function ExcelRange:deleteRows()
    self._com.EntireRow:Delete()
    return self
end

---插入列 (向右移位)
---@return ExcelRange
function ExcelRange:insertColumns()
    self._com.EntireColumn:Insert()
    return self
end

---删除列
---@return ExcelRange
function ExcelRange:deleteColumns()
    self._com.EntireColumn:Delete()
    return self
end

-- ============================================================
-- ExcelSheet 方法
-- ============================================================

---@param com_sheet userdata COM Worksheet 对象
---@return ExcelSheet
function ExcelSheet:new(com_sheet)
    local obj = { _com = com_sheet }
    setmetatable(obj, self)
    return obj
end

---获取工作表名称
---@return string
function ExcelSheet:name()
    return self._com.Name
end

---设置工作表名称
---@param n string
---@return ExcelSheet
function ExcelSheet:setName(n)
    self._com.Name = n
    return self
end

---获取工作表索引 (1-based)
---@return integer
function ExcelSheet:index()
    return self._com.Index
end

---激活该工作表
---@return ExcelSheet
function ExcelSheet:activate()
    self._com:Activate()
    return self
end

---获取/设置单个单元格的值
---
---```lua
---local v = sheet:cell(1, 1)       -- 获取 A1
---sheet:cell(2, 3, "Hello")        -- 设置 C2 = "Hello"
---```
---
---@param r integer 行
---@param c integer 列
---@param v any|nil 值
---@return any|ExcelSheet
function ExcelSheet:cell(r, c, v)
    if v ~= nil then
        self._com.Cells(r, c).Value2 = v
        return self
    else
        return self._com.Cells(r, c).Value2
    end
end

---获取区域对象
---
---```lua
---local rng = sheet:range("A1:C10")
---local rng = sheet:range(1, 1, 10, 3)  -- 等价
---```
---
---@overload fun(addr:string):ExcelRange
---@overload fun(r1:integer, c1:integer, r2:integer, c2:integer):ExcelRange
---@return ExcelRange
function ExcelSheet:range(a, b, c, d)
    if type(a) == "string" then
        return ExcelRange:new(self._com:Range(a))
    else
        return ExcelRange:new(self._com:Range(
            self._com.Cells(a, b),
            self._com.Cells(c, d)
        ))
    end
end

---获取已使用区域
---@return ExcelRange
function ExcelSheet:usedRange()
    return ExcelRange:new(self._com.UsedRange)
end

---获取总行数 (已使用区域)
---@return integer
function ExcelSheet:usedRows()
    return self._com.UsedRange.Rows.Count
end

---获取总列数 (已使用区域)
---@return integer
function ExcelSheet:usedColumns()
    return self._com.UsedRange.Columns.Count
end

---删除该工作表
function ExcelSheet:delete()
    self._com:Delete()
end

---显示/隐藏工作表
---@param visible boolean
---@return ExcelSheet
function ExcelSheet:visible(visible)
    if visible then
        self._com.Visible = -1  -- xlSheetVisible
    else
        self._com.Visible = 0   -- xlSheetHidden
    end
    return self
end

---设置工作表保护
---@param password string|nil
---@return ExcelSheet
function ExcelSheet:protect(password)
    if password then
        self._com:Protect(password)
    else
        self._com:Protect()
    end
    return self
end

---取消保护
---@param password string|nil
---@return ExcelSheet
function ExcelSheet:unprotect(password)
    if password then
        self._com:Unprotect(password)
    else
        self._com:Unprotect()
    end
    return self
end

-- ============================================================
-- ExcelWorkbook 方法
-- ============================================================

---@param com_wb userdata COM Workbook 对象
---@return ExcelWorkbook
function ExcelWorkbook:new(com_wb)
    local obj = { _com = com_wb }
    setmetatable(obj, self)
    return obj
end

---获取工作簿文件路径
---@return string
function ExcelWorkbook:path()
    return self._com.FullName
end

---获取工作簿文件名
---@return string
function ExcelWorkbook:name()
    return self._com.Name
end

---获取工作表数量
---@return integer
function ExcelWorkbook:sheetCount()
    return self._com.Sheets.Count
end

---获取指定名称的工作表
---@param name string
---@return ExcelSheet
function ExcelWorkbook:sheet(name)
    return ExcelSheet:new(self._com.Sheets(name))
end

---获取指定索引的工作表 (1-based)
---@param index integer
---@return ExcelSheet
function ExcelWorkbook:sheetByIndex(index)
    return ExcelSheet:new(self._com.Sheets(index))
end

---获取所有工作表名称列表
---@return string[]
function ExcelWorkbook:sheetNames()
    local names = {}
    for i = 1, self._com.Sheets.Count do
        names[i] = self._com.Sheets(i).Name
    end
    return names
end

---添加工作表
---
---```lua
---local s = wb:addSheet("新表")           -- 最前面插入
---local s = wb:addSheet("新表", nil)      -- 最前面
---local s = wb:addSheet("新表", wb:sheet("Sheet1"))  -- 在 Sheet1 之后
---```
---
---@param name string|nil 名称
---@param after ExcelSheet|nil 插入在某表之后 (nil 则最前)
---@return ExcelSheet
function ExcelWorkbook:addSheet(name, after)
    local com_after = after and after._com or nil
    local com_sheet = self._com.Sheets:Add(com_after)
    if name then
        com_sheet.Name = name
    end
    return ExcelSheet:new(com_sheet)
end

---删除指定名称的工作表
---@param name string
function ExcelWorkbook:deleteSheet(name)
    self._com.Sheets(name):Delete()
end

---保存
function ExcelWorkbook:save()
    self._com:Save()
end

---另存为
---
---支持格式: xlsx, xlsm, xls, csv, pdf
---
---```lua
---wb:saveAs("C:\\report.xlsx")
---wb:saveAs("C:\\report.csv", "csv")
---```
---
---@param path string 文件路径
---@param format string|nil 格式 ("xlsx"|"xlsm"|"xls"|"csv"|"pdf")
function ExcelWorkbook:saveAs(path, format)
    local fileFormat = {
        xlsx = 51,   -- xlOpenXMLWorkbook
        xlsm = 52,   -- xlOpenXMLWorkbookMacroEnabled
        xls  = 56,   -- xlExcel8
        csv  = 6,    -- xlCSV
        pdf  = 57,   -- xlPDF
    }
    local fmt = format and fileFormat[format:lower()] or fileFormat.xlsx
    if fmt then
        self._com.SaveAs(path, fmt)
    else
        self._com.SaveAs(path)
    end
end

---关闭工作簿
---@param save_changes boolean|nil 是否保存更改
function ExcelWorkbook:close(save_changes)
    if save_changes == nil then save_changes = true end
    self._com:Close(save_changes)
end

---设置计算模式
---
---```lua
---wb:calculation("manual")   -- 手动计算
---wb:calculation("auto")     -- 自动计算 (默认)
---```
---
---@param mode string "auto" | "manual"
function ExcelWorkbook:calculation(mode)
    if mode == "manual" then
        self._com.Application.Calculation = -4135  -- xlCalculationManual
    else
        self._com.Application.Calculation = -4105  -- xlCalculationAutomatic
    end
end

---重新计算所有工作表
function ExcelWorkbook:calculate()
    self._com.Application:Calculate()
end

-- ============================================================
-- ExcelApp 方法
-- ============================================================

---创建或连接 Excel 应用实例
---
---若 Excel 已在运行则复用，否则启动新进程。
---
---```lua
---local excel = ExcelApp:new()          -- 不可见
---local excel = ExcelApp:new(true)      -- 可见 (调试用)
---local excel = ExcelApp:new(false)     -- 不可见
---```
---
---@param visible boolean|nil 是否显示 Excel 窗口 (默认 false)
---@return ExcelApp
function ExcelApp:new(visible)
    local luacom = require("luacom")
    local com_excel = luacom.CreateObject("Excel.Application")
    if com_excel == nil then
        return nil, "无法创建 Excel.Application COM 对象"
    end
    com_excel.Visible = visible and true or false
    com_excel.DisplayAlerts = false  -- 不弹出确认对话框
    com_excel.ScreenUpdating = false -- 默认关闭屏幕刷新以提高性能
    local obj = {
        _com = com_excel,
        _luacom = luacom,
    }
    setmetatable(obj, self)
    return obj
end

---设置 Excel 窗口可见性
---@param v boolean
---@return ExcelApp
function ExcelApp:visible(v)
    self._com.Visible = v
    return self
end

---设置是否弹出警告对话框
---@param v boolean
---@return ExcelApp
function ExcelApp:displayAlerts(v)
    self._com.DisplayAlerts = v
    return self
end

---设置是否刷新屏幕 (关闭可提升写入性能)
---@param v boolean
---@return ExcelApp
function ExcelApp:screenUpdating(v)
    self._com.ScreenUpdating = v
    return self
end

---打开已有工作簿
---@param path string 文件路径
---@return ExcelWorkbook
function ExcelApp:open(path)
    local com_wb = self._com.Workbooks:Open(path)
    return ExcelWorkbook:new(com_wb)
end

---新建工作簿
---@return ExcelWorkbook
function ExcelApp:addWorkbook()
    local com_wb = self._com.Workbooks:Add()
    return ExcelWorkbook:new(com_wb)
end

---获取已打开的工作簿数量
---@return integer
function ExcelApp:workbookCount()
    return self._com.Workbooks.Count
end

---获取指定索引的工作簿 (1-based)
---@param index integer
---@return ExcelWorkbook
function ExcelApp:workbook(index)
    return ExcelWorkbook:new(self._com.Workbooks(index))
end

---获取 Excel 版本号
---@return string
function ExcelApp:version()
    return self._com.Version
end

---设置计算模式 (全局)
---
---```lua
---excel:calculation("manual")  -- 手动
---excel:calculation("auto")    -- 自动
---```
---
---@param mode string
---@return ExcelApp
function ExcelApp:calculation(mode)
    if mode == "manual" then
        self._com.Calculation = -4135
    else
        self._com.Calculation = -4105
    end
    return self
end

---退出 Excel 应用程序
---
---退出前自动关闭所有工作簿。
---
---```lua
---excel:quit()         -- 不保存，直接退出
---excel:quit(true)     -- 逐个工作簿保存后退出
---```
---
---@param save_all boolean|nil
function ExcelApp:quit(save_all)
    for i = self._com.Workbooks.Count, 1, -1 do
        local wb = self._com.Workbooks(i)
        if save_all then
            pcall(function() wb:Save() end)
        end
        pcall(function() wb:Close(false) end)  -- false = 不保存
    end
    self._com:Quit()
end

-- ============================================================
-- 便捷函数: 列号 ↔ 字母转换
-- ============================================================

---将列号转为字母 (1→A, 27→AA)
---@param n integer
---@return string
function ExcelApp.columnLetter(n)
    local letters = {}
    while n > 0 do
        local m = (n - 1) % 26
        table.insert(letters, 1, string.char(65 + m))
        n = (n - m) // 26
    end
    return table.concat(letters)
end

---将列字母转为列号 (A→1, AA→27)
---@param s string
---@return integer
function ExcelApp.columnNumber(s)
    local n = 0
    for i = 1, #s do
        n = n * 26 + (string.byte(s, i) - 64)
    end
    return n
end

---生成单元格地址字符串，如 "A1", "AA10"
---@param r integer 行
---@param c integer 列
---@return string
function ExcelApp.cellAddr(r, c)
    return ExcelApp.columnLetter(c) .. tostring(r)
end

---生成区域地址字符串，如 "A1:C10"
---@param r1 integer
---@param c1 integer
---@param r2 integer
---@param c2 integer
---@return string
function ExcelApp.rangeAddr(r1, c1, r2, c2)
    return ExcelApp.cellAddr(r1, c1) .. ":" .. ExcelApp.cellAddr(r2, c2)
end

---将 os.date("*t") 格式的日期 table 转为 Excel 日期序列号
---
---```lua
---local serial = excel.dateSerial(os.date("*t"))
---sheet:cell(1, 1, serial)
---sheet:range("A1"):numberFormat("yyyy-mm-dd")
---```
---
---@param t table {year, month, day, hour, min, sec}
---@return number
function ExcelApp.dateSerial(t)
    local y, m, d = t.year or 1899, t.month or 12, t.day or 30
    if m <= 2 then y = y - 1; m = m + 12 end
    local a = math.floor(y / 100)
    local b = 2 - a + math.floor(a / 4)
    local jd = math.floor(365.25 * (y + 4716)) + math.floor(30.6001 * (m + 1)) + d + b - 1524
    local serial = jd - 2415019  -- 儒略日 → Excel 日期 (基准 1899-12-30)
    -- 加上时间部分
    if t.hour then
        serial = serial + (t.hour * 3600 + (t.min or 0) * 60 + (t.sec or 0)) / 86400
    end
    return serial
end

-- ============================================================
-- 编码工具 (UTF-8 ↔ 系统 ANSI)
--
-- Windows 下 Lua 源码通常为 UTF-8 with BOM 或 GBK。
-- 若源码保存为 UTF-8 without BOM，在中文 Windows (CP936) 上
-- Lua 会按 GBK 解析字符串字面量，导致写入 Excel 的中文乱码。
-- 以下函数通过 powershell 实现可靠转码，解决此问题。
--
-- ```lua
-- local enc = require("excel").encoding
-- local name = enc.forExcel("中文表名")  -- 自动判断是否需要转码
-- sheet:setName(name)
-- ```
-- ============================================================

local encoding = {}

---检测当前系统是否为 Windows
---@return boolean
function encoding.isWindows()
    return package.config:sub(1, 1) == "\\"
end

---获取系统 ANSI 代码页号 (仅 Windows)
---中文系统通常返回 936 (GBK)，英文返回 1252
---@return integer|nil
function encoding.getACP()
    if not encoding.isWindows() then return nil end
    local handle = io.popen("chcp 2>nul", "r")
    if not handle then return nil end
    local result = handle:read("*a")
    handle:close()
    local acp = result:match("(%d+)")
    return acp and tonumber(acp) or nil
end

---将 UTF-8 字符串转为系统 ANSI 编码 (通过 powershell)
---@param utf8_str string
---@return string|nil
---@return string|nil 错误信息
function encoding.toACP(utf8_str)
    if not encoding.isWindows() then
        return utf8_str
    end
    local tmp_in = os.tmpname()
    local tmp_out = os.tmpname()
    local f = io.open(tmp_in, "wb")
    if not f then return nil, "cannot create temp file" end
    f:write(utf8_str)
    f:close()
    local cmd = string.format(
        'powershell -NoProfile -Command "[System.IO.File]::WriteAllText(\'%s\', [System.IO.File]::ReadAllText(\'%s\', [System.Text.Encoding]::UTF8)), [System.Text.Encoding]::Default)"',
        tmp_out, tmp_in
    )
    local handle = io.popen(cmd, "r")
    if not handle then
        os.remove(tmp_in)
        return nil, "cannot execute powershell"
    end
    handle:close()
    local f2 = io.open(tmp_out, "rb")
    if not f2 then
        os.remove(tmp_in)
        return nil, "cannot read converted result"
    end
    local result = f2:read("*a")
    f2:close()
    os.remove(tmp_in)
    os.remove(tmp_out)
    return result
end

---将系统 ANSI 字符串转为 UTF-8 (通过 powershell)
---@param ansi_str string
---@return string|nil
function encoding.toUTF8(ansi_str)
    if not encoding.isWindows() then
        return ansi_str
    end
    local tmp_in = os.tmpname()
    local tmp_out = os.tmpname()
    local f = io.open(tmp_in, "wb")
    if not f then return nil end
    f:write(ansi_str)
    f:close()
    local cmd = string.format(
        'powershell -NoProfile -Command "[System.IO.File]::WriteAllText(\'%s\', [System.IO.File]::ReadAllText(\'%s\', [System.Text.Encoding]::Default)), [System.Text.Encoding]::UTF8)"',
        tmp_out, tmp_in
    )
    local handle = io.popen(cmd, "r")
    if not handle then os.remove(tmp_in); return nil end
    handle:close()
    local f2 = io.open(tmp_out, "rb")
    if not f2 then os.remove(tmp_in); return nil end
    local result = f2:read("*a")
    f2:close()
    os.remove(tmp_in)
    os.remove(tmp_out)
    return result
end

---将字符串转为适合写入 Excel 的编码
---
---在中文 Windows 上，LuaCOM 需要 ANSI/GBK 编码的字符串
---才能正确显示在 Excel 中。此函数在中文代码页下自动转码，
---其他环境原样返回。
---
---```lua
---sheet:cell(1, 1, enc.forExcel("中文内容"))
---```
---
---@param str string 源码中的 UTF-8 字符串字面量
---@return string
function encoding.forExcel(str)
    if encoding.isWindows() then
        local acp = encoding.getACP()
        if acp == 936 or acp == 950 or acp == 54936 then
            local converted, err = encoding.toACP(str)
            if converted then return converted end
        end
    end
    return str
end

-- ============================================================
-- 导出
-- ============================================================

return {
    ExcelApp = ExcelApp,
    ExcelWorkbook = ExcelWorkbook,
    ExcelSheet = ExcelSheet,
    ExcelRange = ExcelRange,
    columnLetter = ExcelApp.columnLetter,
    columnNumber = ExcelApp.columnNumber,
    cellAddr = ExcelApp.cellAddr,
    rangeAddr = ExcelApp.rangeAddr,
    dateSerial = ExcelApp.dateSerial,
    encoding = encoding,
}
