--[[
===============================================================================
  Office.lua  —  LuaCOM Excel/Word 操作库
===============================================================================

  依赖: Lua 5.5 + luacom + ffi (Windows), Office 已安装

  -- Excel --
    local office = require("office")
    local app = office.ExcelApp:new()
    local wb = app:addWorkbook()
    local s = wb:sheetByIndex(1)
    s:cell(1, 1, "Hello")
    wb:saveAs("C:\\result.xlsx")
    app:quit()

  -- Word --
    local word = office.WordApp:new()
    local doc = word:addDocument()
    doc:content():text("Hello World"):bold(true)
    doc:saveAs("C:\\result.docx")
    word:quit()

  编码说明:
    中文 Windows 上 LuaCOM 需要 ANSI 编码字符串。
    1. 保存 .lua 为 UTF-8 with BOM (推荐)
    2. 或使用 encoding.forExcel("中文") / encoding.forExcel("表名")

  对象层级:
    Excel: ExcelApp → ExcelWorkbook → ExcelSheet → ExcelRange
    Word:  WordApp  → WordDocument  → WordRange

===============================================================================
]]

local luacom

-- ============================================================================
--  COM 常量
-- ============================================================================

--- 对齐方式
local XlHAlign = {
    General = 1, Left = -4131, Center = -4108, Right = -4152,
    Fill = 5, Justify = -4130, CenterAcrossSelection = 7, Distributed = -4117,
}
local XlVAlign = {
    Top = -4160, Center = -4108, Bottom = -4107,
    Justify = -4130, Distributed = -4117,
}

--- 边框线型
local XlLineStyle = {
    Continuous = 1, Dash = -4115, Dot = -4118, Double = -4119,
    None = -4142, SlantDashDot = 13,
}

--- 填充图案
local XlPattern = {
    Solid = 1, None = -4142, Gray75 = 12, Gray50 = 13, Gray25 = 14,
}

--- 文件格式 (saveAs)
local XlFileFormat = {
    xlsx = 51, xlsm = 52, xls = 56, csv = 6, pdf = 57,
}

--- 计算模式
local XlCalcManual = -4135   -- xlCalculationManual
local XlCalcAuto   = -4105   -- xlCalculationAutomatic

-- ============================================================================
--  类声明 (必须先声明 local 再定义方法, 避免闭包捕获到全局变量)
-- ============================================================================
local ExcelApp = {}
ExcelApp.__index = ExcelApp

local ExcelWorkbook = {}
ExcelWorkbook.__index = ExcelWorkbook

local ExcelSheet = {}
ExcelSheet.__index = ExcelSheet

local ExcelRange = {}
ExcelRange.__index = ExcelRange

local WordApp = {}
WordApp.__index = WordApp

local WordDocument = {}
WordDocument.__index = WordDocument

local WordRange = {}
WordRange.__index = WordRange

-- ============================================================================
--  内部: 创建 Lua wrapper 对象的工厂
-- ============================================================================

local function markReleased(wrapper)
    wrapper._released = true
end

local function isReleased(wrapper)
    return wrapper._released == true
end

local function newWrapper(cls, com_obj)
    return setmetatable({ _com = com_obj }, cls)
end

-- ============================================================================
--  Class: ExcelApp  —  Excel 应用程序
-- ============================================================================

--- GC 终结器: 对象被回收时自动关闭 Excel (阻止僵尸进程)
function ExcelApp.__gc(app)
    if not isReleased(app) then
        pcall(function() ExcelApp.quit(app) end)
    end
end

--- to-be-closed 支持: Lua 5.5 `local app <close> = ExcelApp:new()`
--- 变量离开作用域时自动调用 (优于 __gc — 确定性释放)
function ExcelApp.__close(app, _err)
    if not isReleased(app) then
        pcall(function() ExcelApp.quit(app) end)
    end
end

--- 创建/连接 Excel 实例
---
--- 支持自动资源管理:
---   GC: 对象被回收时自动调 quit()  → 防僵尸 Excel 进程
---   to-be-closed: `local app <close> = ExcelApp:new()` → 离开作用域自动释放
---
---```lua
---local app = ExcelApp:new()             -- 不可见 (默认)
---local app = ExcelApp:new(true)         -- 可见 (调试)
---local app <close> = ExcelApp:new()     -- 自动清理 (Lua 5.5)
---```
---
---@param visible boolean|nil 是否显示窗口 (默认 false)
---@return ExcelApp|nil
---@return string|nil 错误信息
function ExcelApp:new(visible)
    luacom = require("luacom")
    local com = luacom.CreateObject("Excel.Application")
    if not com then
        return nil, "cannot create Excel.Application COM object"
    end
    com.Visible = visible and true or false
    com.DisplayAlerts = false   -- 不弹对话框
    com.ScreenUpdating = false  -- 默认不刷新 (提升批量写入性能)
    return newWrapper(ExcelApp, com)
end

--- 窗口可见性
---@param v boolean
---@return ExcelApp
function ExcelApp:visible(v)
    self._com.Visible = v
    return self
end

--- 弹窗开关 (默认关闭)
---@param v boolean
---@return ExcelApp
function ExcelApp:displayAlerts(v)
    self._com.DisplayAlerts = v
    return self
end

--- 屏幕刷新开关 (关闭可提升写入性能)
---@param v boolean
---@return ExcelApp
function ExcelApp:screenUpdating(v)
    self._com.ScreenUpdating = v
    return self
end

--- 打开已有工作簿
---@param path string 文件路径
---@return ExcelWorkbook
function ExcelApp:open(path)
    return newWrapper(ExcelWorkbook, self._com.Workbooks:Open(path))
end

--- 新建空白工作簿
---@return ExcelWorkbook
function ExcelApp:addWorkbook()
    return newWrapper(ExcelWorkbook, self._com.Workbooks:Add())
end

--- 已打开工作簿数量
---@return integer
function ExcelApp:workbookCount()
    return self._com.Workbooks.Count
end

--- 按索引获取工作簿 (1-based)
---@param index integer
---@return ExcelWorkbook
function ExcelApp:workbook(index)
    return newWrapper(ExcelWorkbook, self._com.Workbooks(index))
end

--- Excel 版本号
---@return string
function ExcelApp:version()
    return self._com.Version
end

--- 全局计算模式
---
---```lua
---app:calculation("manual")   -- 手动
---app:calculation("auto")     -- 自动
---```
---
---@param mode string "auto"|"manual"
---@return ExcelApp
function ExcelApp:calculation(mode)
    self._com.Calculation = (mode == "manual") and XlCalcManual or XlCalcAuto
    return self
end

--- 退出 Excel
---
--- 先关闭所有工作簿 (默认不保存)，再退出应用程序。
---
---```lua
---app:quit()       -- 丢弃所有更改
---app:quit(true)   -- 保存所有工作簿后退出
---```
---
---@param save_all boolean|nil
function ExcelApp:quit(save_all)
    if isReleased(self) then return end  -- 已退出，防止重复调用
    markReleased(self)
    -- 关闭所有工作簿 (容错: 每个 wb 独立 pcall, 一个失败不影响其他)
    for i = self._com.Workbooks.Count, 1, -1 do
        local wb = self._com.Workbooks(i)
        if save_all then pcall(function() wb:Save() end) end
        pcall(function() wb:Close(false) end)  -- false = 不保存
    end
    pcall(function() self._com:Quit() end)
end

-- ============================================================================
--  Class: ExcelApp  —  静态工具方法
-- ============================================================================

--- 列号 → 字母   (1→"A", 27→"AA", 702→"ZZ")
---@param n integer
---@return string
function ExcelApp.columnLetter(n)
    local parts = {}
    while n > 0 do
        local m = (n - 1) % 26
        table.insert(parts, 1, string.char(65 + m))
        n = (n - m) // 26
    end
    return table.concat(parts)
end

--- 字母 → 列号   ("A"→1, "AA"→27)
---@param s string
---@return integer
function ExcelApp.columnNumber(s)
    local n = 0
    for i = 1, #s do
        n = n * 26 + (string.byte(s, i) - 64)
    end
    return n
end

--- 单元格地址   (1, 1) → "A1"
---@param r integer 行
---@param c integer 列
---@return string
function ExcelApp.cellAddr(r, c)
    return ExcelApp.columnLetter(c) .. tostring(r)
end

--- 区域地址   (1, 1, 10, 3) → "A1:C10"
---@param r1 integer
---@param c1 integer
---@param r2 integer
---@param c2 integer
---@return string
function ExcelApp.rangeAddr(r1, c1, r2, c2)
    return ExcelApp.cellAddr(r1, c1) .. ":" .. ExcelApp.cellAddr(r2, c2)
end

--- os.date("*t") 日期 table → Excel 日期序列号
---
---```lua
---s:cell(1,1, excel.dateSerial(os.date("*t")))
---s:range("A1"):numberFormat("yyyy-mm-dd")
---```
---
---@param t table {year, month, day, hour?, min?, sec?}
---@return number
function ExcelApp.dateSerial(t)
    local y, m, d = t.year or 1899, t.month or 12, t.day or 30
    if m <= 2 then y = y - 1; m = m + 12 end
    local a = math.floor(y / 100)
    local b = 2 - a + math.floor(a / 4)
    local jd = math.floor(365.25 * (y + 4716))
             + math.floor(30.6001 * (m + 1)) + d + b - 1524
    local serial = jd - 2415019   -- 儒略日 → Excel 纪元 (1899-12-30)
    if t.hour then
        serial = serial + (t.hour * 3600 + (t.min or 0) * 60 + (t.sec or 0)) / 86400
    end
    return serial
end


-- ============================================================================
--  Class: ExcelWorkbook  —  工作簿
-- ============================================================================

function ExcelWorkbook:new(com_wb)
    return newWrapper(ExcelWorkbook, com_wb)
end

--- 完整路径
---@return string
function ExcelWorkbook:path()
    return self._com.FullName
end

--- 文件名 (含扩展名)
---@return string
function ExcelWorkbook:name()
    return self._com.Name
end

-- ---- 工作表管理 ----

--- 工作表数量
---@return integer
function ExcelWorkbook:sheetCount()
    return self._com.Sheets.Count
end

--- 按名称获取工作表
---@param name string
---@return ExcelSheet
function ExcelWorkbook:sheet(name)
    return newWrapper(ExcelSheet, self._com.Sheets(name))
end

--- 按索引获取工作表 (1-based)
---@param index integer
---@return ExcelSheet
function ExcelWorkbook:sheetByIndex(index)
    return newWrapper(ExcelSheet, self._com.Sheets(index))
end

--- 所有工作表名称列表
---@return string[]
function ExcelWorkbook:sheetNames()
    local names = {}
    for i = 1, self._com.Sheets.Count do
        names[i] = self._com.Sheets(i).Name
    end
    return names
end

--- 添加工作表
---
---```lua
---wb:addSheet("新表")                              -- 最前
---wb:addSheet("新表", wb:sheet("Sheet1"))           -- 在 Sheet1 之后
---```
---
---@param name string|nil
---@param after ExcelSheet|nil 插入位置 (nil = 最前)
---@return ExcelSheet
function ExcelWorkbook:addSheet(name, after)
    local com_after = after and after._com or nil
    local com_sheet = self._com.Sheets:Add(com_after)
    if name then com_sheet.Name = name end
    return newWrapper(ExcelSheet, com_sheet)
end

--- 按名称删除工作表
---@param name string
function ExcelWorkbook:deleteSheet(name)
    self._com.Sheets(name):Delete()
end

-- ---- 保存 / 关闭 ----

--- 保存
function ExcelWorkbook:save()
    self._com:Save()
end

--- 另存为
---
--- 支持格式: xlsx (默认), xlsm, xls, csv, pdf
---
---```lua
---wb:saveAs("C:\\report.xlsx")
---wb:saveAs("C:\\data.csv", "csv")
---```
---
---@param path string
---@param format string|nil "xlsx"|"xlsm"|"xls"|"csv"|"pdf"
function ExcelWorkbook:saveAs(path, format)
    local fmt = format and XlFileFormat[format:lower()] or XlFileFormat.xlsx
    self._com:SaveAs(path, fmt or XlFileFormat.xlsx)
end

--- 关闭工作簿
---@param save_changes boolean|nil 默认 true
function ExcelWorkbook:close(save_changes)
    if save_changes == nil then save_changes = true end
    self._com:Close(save_changes)
end

-- ---- 计算 ----

--- 设置计算模式
---
---```lua
---wb:calculation("manual")   -- 手动
---wb:calculation("auto")     -- 自动 (默认)
---```
---
---@param mode string "auto"|"manual"
function ExcelWorkbook:calculation(mode)
    self._com.Application.Calculation = (mode == "manual") and XlCalcManual or XlCalcAuto
end

--- 强制重新计算
function ExcelWorkbook:calculate()
    self._com.Application:Calculate()
end


-- ============================================================================
--  Class: ExcelSheet  —  工作表
-- ============================================================================

function ExcelSheet:new(com_sheet)
    return newWrapper(ExcelSheet, com_sheet)
end

--- 名称
---@return string
function ExcelSheet:name()
    return self._com.Name
end

--- 改名
---@param n string
---@return ExcelSheet
function ExcelSheet:setName(n)
    self._com.Name = n
    return self
end

--- 索引 (1-based)
---@return integer
function ExcelSheet:index()
    return self._com.Index
end

--- 激活 (设为当前工作表)
---@return ExcelSheet
function ExcelSheet:activate()
    self._com:Activate()
    return self
end

--- 删除本工作表
function ExcelSheet:delete()
    self._com:Delete()
end

--- 显示/隐藏
---@param v boolean
---@return ExcelSheet
function ExcelSheet:visible(v)
    self._com.Visible = v and -1 or 0   -- -1=xlSheetVisible, 0=xlSheetHidden
    return self
end

--- 保护工作表
---@param password string|nil
---@return ExcelSheet
function ExcelSheet:protect(password)
    if password then self._com:Protect(password) else self._com:Protect() end
    return self
end

--- 取消保护
---@param password string|nil
---@return ExcelSheet
function ExcelSheet:unprotect(password)
    if password then self._com:Unprotect(password) else self._com:Unprotect() end
    return self
end

-- ---- 单元格 / 区域 ----

--- 获取或设置单个单元格的值
---
--- **注意**: 此方法返回单元格值或 Sheet，**不返回 Range**。
--- 如需链式格式化单元格，请用 `range()` 获取 ExcelRange：
---   s:range("A1"):bold(true):halign("Center")
---   s:range(1, 1, 1, 1):value("Hello"):bold(true)
---
---```lua
---local v = s:cell(1, 1)          -- 读 A1 (返回值)
---s:cell(2, 3, "Hello")           -- 写 C2 (返回 Sheet)
---```
---
---@param r integer 行
---@param c integer 列
---@param v any|nil 值 (不传 = 读取)
---@return any|ExcelSheet
function ExcelSheet:cell(r, c, v)
    if v ~= nil then
        self._com.Cells(r, c).Value2 = v
        return self
    else
        return self._com.Cells(r, c).Value2
    end
end

--- 获取区域对象
---
---```lua
---local rng = s:range("A1:C10")
---local rng = s:range(1, 1, 10, 3)  -- 等价
---```
---
---@overload fun(addr:string):ExcelRange
---@overload fun(r1:integer, c1:integer, r2:integer, c2:integer):ExcelRange
---@return ExcelRange
function ExcelSheet:range(a, b, c, d)
    if type(a) == "string" then
        return newWrapper(ExcelRange, self._com:Range(a))
    end
    return newWrapper(ExcelRange, self._com:Range(
        self._com.Cells(a, b),
        self._com.Cells(c, d)))
end

--- 已使用区域
---@return ExcelRange
function ExcelSheet:usedRange()
    return newWrapper(ExcelRange, self._com.UsedRange)
end

--- 已使用行数
---@return integer
function ExcelSheet:usedRows()
    return self._com.UsedRange.Rows.Count
end

--- 已使用列数
---@return integer
function ExcelSheet:usedColumns()
    return self._com.UsedRange.Columns.Count
end


-- ============================================================================
--  Class: ExcelRange  —  单元格 / 区域
-- ============================================================================

function ExcelRange:new(com_range)
    return newWrapper(ExcelRange, com_range)
end

-- ---- 值 & 公式 ----

--- 获取或设置值
---
--- 自动处理:
---   nil  → 清空单元格
---   date table (含 .year) → 调用 dateSerial 转为序列号
---
---```lua
---local v = rng:value()       -- 读
---rng:value("Hello")          -- 写
---rng:value(nil)              -- 清空
---rng:value(os.date("*t"))    -- 写日期
---```
---
---@overload fun():any
---@overload fun(v:any):ExcelRange
function ExcelRange:value(v)
    if v == nil then v = "" end
    if v ~= nil then
        if type(v) == "table" and v.year then
            v = ExcelApp.dateSerial(v)
        end
        self._com.Value2 = v
        return self
    end
    return self._com.Value2
end

--- 获取或设置公式
---
---```lua
---rng:formula("=SUM(A1:A10)")
---local f = rng:formula()
---```
---
---@overload fun():string
---@overload fun(f:string):ExcelRange
function ExcelRange:formula(f)
    if f ~= nil then self._com.Formula = f; return self end
    return self._com.Formula
end

-- ---- 尺寸 & 位置 ----

--- 行数
---@return integer
function ExcelRange:rows()    return self._com.Rows.Count       end

--- 列数
---@return integer
function ExcelRange:columns() return self._com.Columns.Count    end

--- 地址 (如 "$A$1:$C$10")
---@return string
function ExcelRange:address() return self._com.Address         end

--- 首行号 (1-based)
---@return integer
function ExcelRange:row()     return self._com.Row             end

--- 首列号 (1-based)
---@return integer
function ExcelRange:column()  return self._com.Column          end

--- 指定偏移单元格
---@param r integer 行偏移
---@param c integer 列偏移
---@return ExcelRange
function ExcelRange:offset(r, c)
    return newWrapper(ExcelRange, self._com.Offset(r, c))
end

--- 指定子单元格 (1-based, 相对于区域)
---@param r integer 行
---@param c integer 列
---@return ExcelRange
function ExcelRange:cell(r, c)
    return newWrapper(ExcelRange, self._com.Cells(r, c))
end

--- 整行
---@return ExcelRange
function ExcelRange:entireRow()
    return newWrapper(ExcelRange, self._com.EntireRow)
end

--- 整列
---@return ExcelRange
function ExcelRange:entireColumn()
    return newWrapper(ExcelRange, self._com.EntireColumn)
end

-- ---- 字体 (Font) ----

--- 字体名称
---@param name string
---@return ExcelRange
function ExcelRange:fontName(name)
    self._com.Font.Name = name; return self
end

--- 字号
---@param size number
---@return ExcelRange
function ExcelRange:fontSize(size)
    self._com.Font.Size = size; return self
end

--- 字体颜色 (RGB)
---@param rgb integer 例: 0xFF0000 = 红色
---@return ExcelRange
function ExcelRange:fontColor(rgb)
    self._com.Font.Color = rgb; return self
end

--- 粗体  (bold()=读, bold(bool)=写)
---@overload fun():boolean
---@overload fun(enable:boolean):ExcelRange
function ExcelRange:bold(enable)
    if enable ~= nil then self._com.Font.Bold = enable; return self end
    return self._com.Font.Bold
end

--- 斜体
---@overload fun():boolean
---@overload fun(enable:boolean):ExcelRange
function ExcelRange:italic(enable)
    if enable ~= nil then self._com.Font.Italic = enable; return self end
    return self._com.Font.Italic
end

--- 下划线
---@overload fun():boolean
---@overload fun(enable:boolean):ExcelRange
function ExcelRange:underline(enable)
    if enable ~= nil then self._com.Font.Underline = enable; return self end
    return self._com.Font.Underline
end

-- ---- 对齐 (Alignment) ----

--- 水平对齐
---
--- 字符串名或整数枚举值均可。
--- 字符串: "General"|"Left"|"Center"|"Right"|"Fill"|"Justify"|"CenterAcrossSelection"|"Distributed"
---
---```lua
---rng:halign("Center")
---rng:halign(-4108)  -- 等价
---```
---
---@param align string|integer
---@return ExcelRange
function ExcelRange:halign(align)
    self._com.HorizontalAlignment = XlHAlign[align] or align
    return self
end

--- 垂直对齐
---
--- 字符串: "Top"|"Center"|"Bottom"|"Justify"|"Distributed"
---
---@param align string|integer
---@return ExcelRange
function ExcelRange:valign(align)
    self._com.VerticalAlignment = XlVAlign[align] or align
    return self
end

--- 自动换行
---@param enable boolean
---@return ExcelRange
function ExcelRange:wrapText(enable)
    self._com.WrapText = enable; return self
end

-- ---- 背景 (Interior) ----

--- 背景颜色 (RGB)
---@param rgb integer
---@return ExcelRange
function ExcelRange:bgColor(rgb)
    self._com.Interior.Color = rgb; return self
end

--- 填充图案
---
--- 字符串: "Solid"|"None"|"Gray75"|"Gray50"|"Gray25"
---
---@param pattern string|integer
---@return ExcelRange
function ExcelRange:bgPattern(pattern)
    self._com.Interior.Pattern = XlPattern[pattern] or pattern
    return self
end

-- ---- 边框 (Borders) ----

--- 单独边框
---
--- border_index: 7=左, 8=上, 9=下, 10=右, 11=内部竖线, 12=内部横线
--- line_style: "Continuous"|"Dash"|"Dot"|"Double"|"None"|"SlantDashDot" 或整数
---
---```lua
---rng:border(7, "Continuous")   -- 左边框
---rng:border(8, 1)              -- 上边框 (1=Continuous)
---```
---
---@param border_index integer
---@param line_style string|integer
---@return ExcelRange
function ExcelRange:border(border_index, line_style)
    self._com.Borders(border_index).LineStyle = XlLineStyle[line_style] or line_style
    return self
end

--- 全部边框 (含内部线)
---@param line_style string|integer
---@return ExcelRange
function ExcelRange:borderAll(line_style)
    local ls = XlLineStyle[line_style] or line_style
    for i = 7, 12 do self._com.Borders(i).LineStyle = ls end
    return self
end

--- 外边框 (不含内部线)
---@param line_style string|integer
---@return ExcelRange
function ExcelRange:borderOutline(line_style)
    local ls = XlLineStyle[line_style] or line_style
    for i = 7, 10 do self._com.Borders(i).LineStyle = ls end
    return self
end

-- ---- 数字格式 ----

--- 数字格式
---
---```lua
---rng:numberFormat("0.00")        -- 两位小数
---rng:numberFormat("#,##0.00")    -- 千分位
---rng:numberFormat("yyyy-mm-dd")  -- 日期
---rng:numberFormat("@")           -- 文本
---```
---
---@param fmt string
---@return ExcelRange
function ExcelRange:numberFormat(fmt)
    self._com.NumberFormat = fmt; return self
end

-- ---- 列宽 / 行高 ----

--- 列宽
---@param w number
---@return ExcelRange
function ExcelRange:columnWidth(w)
    self._com.ColumnWidth = w; return self
end

--- 自适应列宽
---@return ExcelRange
function ExcelRange:autoFit()
    self._com.EntireColumn:AutoFit(); return self
end

--- 行高
---@param h number
---@return ExcelRange
function ExcelRange:rowHeight(h)
    self._com.RowHeight = h; return self
end

-- ---- 合并 ----

--- 合并单元格
---@return ExcelRange
function ExcelRange:merge()
    self._com:Merge(); return self
end

--- 取消合并
---@return ExcelRange
function ExcelRange:unmerge()
    self._com:UnMerge(); return self
end

-- ---- 复制 / 剪切 / 粘贴 ----

function ExcelRange:copy()           self._com:Copy(); return self end
function ExcelRange:cut()            self._com:Cut();  return self end

--- 粘贴到目标区域
---@param target ExcelRange
---@return ExcelRange
function ExcelRange:pasteTo(target)
    self._com:Copy(target._com); return self
end

-- ---- 清除 ----

--- 清除内容 (保留格式)
function ExcelRange:clear()          self._com:ClearContents(); end

--- 全部清除 (含格式)
function ExcelRange:clearAll()       self._com:Clear();         end

-- ---- 插入 / 删除 ----

function ExcelRange:insertRows()     self._com.EntireRow:Insert();    end
function ExcelRange:deleteRows()     self._com.EntireRow:Delete();    end
function ExcelRange:insertColumns()  self._com.EntireColumn:Insert(); end
function ExcelRange:deleteColumns()  self._com.EntireColumn:Delete(); end


-- ============================================================================
--  Module: encoding  —  UTF-8 ↔ 系统 ANSI (GBK) 转码
-- ============================================================================
--  在中文 Windows (CP936) 上，LuaCOM 需要 ANSI/GBK 编码的字符串才能
--  正确写入 Excel。若源文件为 UTF-8 without BOM，中文会出现乱码。
--  本模块通过 FFI 调用 Win32 API (MultiByteToWideChar / WideCharToMultiByte)
--  实现编码转换，无需 powershell、临时文件或额外 COM 对象。
--
--  使用:
--    local enc = require("excel").encoding
--    sheet:cell(1, 1, enc.forExcel("中文"))  -- 自动转码
--    sheet:setName(enc.forExcel("表名"))
--
--  最佳实践: 将 .lua 文件保存为 UTF-8 with BOM，则无需调用 forExcel()
-- ============================================================================
local encoding = {}

function encoding.isWindows()
    return package.config:sub(1, 1) == "\\"
end

function encoding.getACP()
    if not encoding.isWindows() then return nil end
    local h = io.popen("chcp 2>nul", "r")
    if not h then return nil end
    local r = h:read("*a"); h:close()
    return tonumber((r:match("(%d+)")))
end

-- 延迟加载 FFI (仅 Windows 需要)
local _ffi_ready = false
local function _initFFI()
    if _ffi_ready then return true end
    local ok, ffi = pcall(require, "ffi")
    if not ok then return false end
    pcall(function()
        ffi.cdef([[
            typedef unsigned short wchar_t;
            int MultiByteToWideChar(unsigned int, unsigned long,
                const char*, int, wchar_t*, int);
            int WideCharToMultiByte(unsigned int, unsigned long,
                const wchar_t*, int, char*, int,
                const char*, int*);
        ]])
    end)
    _ffi_ready = true
    return true
end
local function _ffi() return require("ffi") end

--- UTF-8 → 系统 ANSI (Win32 MultiByteToWideChar → WideCharToMultiByte)
function encoding.toACP(utf8_str)
    if not encoding.isWindows() then return utf8_str end
    if #utf8_str == 0 then return utf8_str end
    if not _initFFI() then return nil end
    local ffi = _ffi()
    local ok, result = pcall(function()
        local C, src = ffi.C, utf8_str
        local wlen = C.MultiByteToWideChar(65001, 0, src, #src, nil, 0)
        if wlen == 0 then return nil end
        local wbuf = ffi.new("wchar_t[?]", wlen + 1)
        C.MultiByteToWideChar(65001, 0, src, #src, wbuf, wlen)
        local alen = C.WideCharToMultiByte(0, 0, wbuf, wlen, nil, 0, nil, nil)
        if alen == 0 then return nil end
        local abuf = ffi.new("char[?]", alen + 1)
        C.WideCharToMultiByte(0, 0, wbuf, wlen, abuf, alen, nil, nil)
        return ffi.string(abuf, alen)
    end)
    if not ok then return nil else return result end
end

--- 系统 ANSI → UTF-8 (Win32 MultiByteToWideChar → WideCharToMultiByte)
function encoding.toUTF8(ansi_str)
    if not encoding.isWindows() then return ansi_str end
    if #ansi_str == 0 then return ansi_str end
    if not _initFFI() then return nil end
    local ffi = _ffi()
    local ok, result = pcall(function()
        local C, src = ffi.C, ansi_str
        local wlen = C.MultiByteToWideChar(0, 0, src, #src, nil, 0)
        if wlen == 0 then return nil end
        local wbuf = ffi.new("wchar_t[?]", wlen + 1)
        C.MultiByteToWideChar(0, 0, src, #src, wbuf, wlen)
        local ulen = C.WideCharToMultiByte(65001, 0, wbuf, wlen, nil, 0, nil, nil)
        if ulen == 0 then return nil end
        local ubuf = ffi.new("char[?]", ulen + 1)
        C.WideCharToMultiByte(65001, 0, wbuf, wlen, ubuf, ulen, nil, nil)
        return ffi.string(ubuf, ulen)
    end)
    if not ok then return nil else return result end
end

--- 为写入 Excel 准备字符串
---
--- 中文代码页 (936/950/54936) 下自动 UTF8→ANSI 转码，
--- 其他环境原样返回。
---
---@param str string 源码 UTF-8 字符串
---@return string
function encoding.forExcel(str)
    if encoding.isWindows() then
        local acp = encoding.getACP()
        if acp == 936 or acp == 950 or acp == 54936 then
            local converted = encoding.toACP(str)
            if converted then return converted end
        end
    end
    return str
end


-- ============================================================================
--  导出
-- ============================================================================
-- ============================================================================
--  Class: WordApp  —  Word 应用程序
-- ============================================================================

function WordApp.__gc(app)
    if not isReleased(app) then
        pcall(function() WordApp.quit(app) end)
    end
end

function WordApp.__close(app, _err)
    if not isReleased(app) then
        pcall(function() WordApp.quit(app) end)
    end
end

--- 创建 Word 实例
---
---```lua
---local word = office.WordApp:new()       -- 不可见
---local word = office.WordApp:new(true)   -- 可见 (调试)
---local word <close> = office.WordApp:new() -- 自动清理
---```
---@param visible boolean|nil
---@return WordApp|nil
function WordApp:new(visible)
    luacom = require("luacom")
    local com = luacom.CreateObject("Word.Application")
    if not com then return nil end
    com.Visible = visible and true or false
    com.DisplayAlerts = false
    com.ScreenUpdating = false
    return newWrapper(WordApp, com)
end

function WordApp:visible(v)    self._com.Visible = v; return self end
function WordApp:displayAlerts(v) self._com.DisplayAlerts = v; return self end
function WordApp:screenUpdating(v) self._com.ScreenUpdating = v; return self end

--- 打开已有文档
---@return WordDocument
function WordApp:open(path)
    return newWrapper(WordDocument, self._com.Documents:Open(path))
end

--- 新建空白文档
---@return WordDocument
function WordApp:addDocument()
    return newWrapper(WordDocument, self._com.Documents:Add())
end

function WordApp:documentCount() return self._com.Documents.Count end

--- 按索引获取文档 (1-based)
---@return WordDocument
function WordApp:document(index)
    return newWrapper(WordDocument, self._com.Documents(index))
end

function WordApp:version() return self._com.Version end

--- 退出 Word
---@param save_all boolean|nil
function WordApp:quit(save_all)
    if isReleased(self) then return end
    markReleased(self)
    for i = self._com.Documents.Count, 1, -1 do
        local doc = self._com.Documents(i)
        if save_all then pcall(function() doc:Save() end) end
        pcall(function() doc:Close(false) end)
    end
    pcall(function() self._com:Quit() end)
end

-- ============================================================================
--  Class: WordDocument  —  Word 文档
-- ============================================================================

function WordDocument:new(com_doc)
    return newWrapper(WordDocument, com_doc)
end

function WordDocument:name()     return self._com.Name end
function WordDocument:path()     return self._com.FullName end

--- 获取文档 Range (0=开头, -1=结尾, 省略则全文)
---@param start integer|nil
---@param end_ integer|nil
---@return WordRange
function WordDocument:range(start, end_)
    local s = start or 0
    if not end_ then
        -- Content.End 返回文档末尾位置
        end_ = self._com.Content.End_
    end
    return newWrapper(WordRange, self._com:Range(s, end_))
end

--- 全文 Range
---@return WordRange
function WordDocument:content()
    return newWrapper(WordRange, self._com.Content)
end

--- 段落数
function WordDocument:paragraphCount()
    return self._com.Paragraphs.Count
end

--- 表格数
function WordDocument:tableCount()
    return self._com.Tables.Count
end

function WordDocument:save()     self._com:Save() end
function WordDocument:saveAs(path) self._com:SaveAs(path) end

---@param save_changes boolean|nil
function WordDocument:close(save_changes)
    if save_changes == nil then save_changes = true end
    self._com:Close(save_changes)
end

-- ============================================================================
--  Class: WordRange  —  Word 文档区域
-- ============================================================================

function WordRange:new(com_range)
    return newWrapper(WordRange, com_range)
end

--- 获取或设置文本
---@overload fun():string
---@overload fun(v:string):WordRange
function WordRange:text(v)
    if v ~= nil then self._com.Text = v; return self end
    return self._com.Text
end

-- 字体
---@overload fun():boolean
---@overload fun(v:boolean):WordRange
function WordRange:bold(v)
    if v ~= nil then self._com.Font.Bold = v; return self end
    return self._com.Font.Bold
end

---@overload fun():boolean
---@overload fun(v:boolean):WordRange
function WordRange:italic(v)
    if v ~= nil then self._com.Font.Italic = v; return self end
    return self._com.Font.Italic
end

function WordRange:fontName(name) self._com.Font.Name = name; return self end
function WordRange:fontSize(size) self._com.Font.Size = size; return self end
function WordRange:fontColor(rgb) self._com.Font.Color = rgb; return self end

--- 段落对齐: "Left"|"Center"|"Right"|"Justify"
function WordRange:halign(align)
    local map = { Left = 0, Center = 1, Right = 2, Justify = 3 }
    self._com.ParagraphFormat.Alignment = map[align] or align
    return self
end

--- 复制 / 粘贴
function WordRange:copy() self._com:Copy(); return self end
function WordRange:paste() self._com:Paste(); return self end

--- 插入段落
---@return WordRange
function WordRange:insertParagraph()
    return newWrapper(WordRange, self._com:InsertParagraphAfter())
end

--- 插入表格
---@param rows integer
---@param cols integer
---@return WordRange (table 区域)
function WordRange:insertTable(rows, cols)
    local tbl = self._com.Tables:Add(self._com, rows, cols)
    return newWrapper(WordRange, tbl.Range)
end

-- ============================================================================
--  导出
-- ============================================================================
return {
    -- Excel
    ExcelApp      = ExcelApp,
    ExcelWorkbook = ExcelWorkbook,
    ExcelSheet    = ExcelSheet,
    ExcelRange    = ExcelRange,
    -- Word
    WordApp       = WordApp,
    WordDocument  = WordDocument,
    WordRange     = WordRange,
    -- 工具
    columnLetter  = ExcelApp.columnLetter,
    columnNumber  = ExcelApp.columnNumber,
    cellAddr      = ExcelApp.cellAddr,
    rangeAddr     = ExcelApp.rangeAddr,
    dateSerial    = ExcelApp.dateSerial,
    encoding      = encoding,
}
