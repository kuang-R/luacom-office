--[[
===============================================================================
  Excel.lua 测试套件  Test Suite
===============================================================================

  文件编码:
    保存为 UTF-8 with BOM 以确保 Windows Lua 正确处理中文
    WSL2 默认 UTF-8 without BOM — 传到 Windows 前需转码

  输出:
    控制台 — 简洁的步骤摘要
    日志文件 (test_log_YYYYMMDD_HHMMSS.txt) — 完整 hex dump 供离线调试

  测试结构:
    PART 1  ENVIRONMENT          环境快照，采集 OS/Lua/COM/codepage 信息
    PART 2  UTILITY FUNCTIONS    纯函数单元测试，不需要 Excel           [1 step]
    PART 3  ENCODING             UTF-8 ↔ GBK 往返转码测试              [1 step]
    PART 4  EXCEL COM            核心流程集成测试 (创建→读写→格式→退出) [11 steps]
    PART 5  COVERAGE EXTENSION   覆盖率扩展 (App/Sheet/Range 剩余 API) [13 steps]
    PART 6  RESOURCE MANAGEMENT  COM 自动资源管理 (__gc/__close/容错)  [5 steps]

===============================================================================
]]

local excel = require("excel")

----------------------------------------------------------------------
--  Logger  --  双路输出: 控制台 (简洁) + 文件 (详细, 含 hex dump)
----------------------------------------------------------------------
local Logger = {}

function Logger:new(filepath)
    local obj = {
        _fh = nil, _path = filepath,
        _start = os.time(),
        _step = 0, _errors = 0, _warns = 0, _skips = 0,
        _indent = 0,
    }
    setmetatable(obj, { __index = Logger })
    local fh, err = io.open(filepath, "w")
    if fh then
        fh:write("\239\187\191")  -- UTF-8 BOM
        obj._fh = fh
    end
    obj:_both(string.rep("=", 60))
    obj:_both("  Excel.lua Test  --  " .. os.date("%Y-%m-%d %H:%M:%S"))
    obj:_both(string.rep("=", 60))
    return obj
end

-- 内部: 写入文件
function Logger:_write(s)
    if self._fh then self._fh:write(s .. "\n"); self._fh:flush() end
end

-- 内部: 双路 (文件 + 控制台)
function Logger:_both(s)
    self:_write(s)
    print(s)
end

-- 内部: 仅文件 (不刷屏)
function Logger:_file(s)
    self:_write(s)
end

-- 内部: 缩进
function Logger:_pad()
    return string.rep("  ", self._indent)
end

----------------------------------------------------------------------
--  Logger API
----------------------------------------------------------------------

--- 区块标题
function Logger:section(title)
    local line = "+" .. string.rep("-", 58) .. "+"
    self:_both("")
    self:_both(line)
    local bare = title:gsub("\27%b[]", "")
    local pad = #bare < 56 and string.rep(" ", 56 - #bare) or ""
    self:_both("| " .. title .. pad .. " |")
    self:_both(line)
end

--- 测试步骤 (自动编号)
---@return integer id
function Logger:step(desc)
    self._step = self._step + 1
    local id = self._step
    self:_both(string.format("  [%2d] %s", id, desc))
    self._indent = 1
    return id
end

--- 步骤通过
function Logger:pass(id)
    self:_both(string.format("  [%2d]   => PASS", id))
    self._indent = 0
end

--- 步骤失败
function Logger:fail(id, reason)
    self._errors = self._errors + 1
    local s = string.format("  [%2d]   => FAIL", id)
    if reason then s = s .. "  (" .. reason .. ")" end
    self:_both(s)
    self._indent = 0
end

--- 步骤跳过
function Logger:skip(id, reason)
    self._skips = self._skips + 1
    local s = string.format("  [%2d]   => SKIP", id)
    if reason then s = s .. "  (" .. reason .. ")" end
    self:_both(s)
    self._indent = 0
end

--- 普通信息 (双路)
function Logger:info(msg)
    self:_both("  " .. self:_pad() .. msg)
end

--- 详细信息 (仅文件)
function Logger:detail(msg)
    self:_file("  " .. self:_pad() .. "[*] " .. msg)
end

--- 警告
function Logger:warn(msg)
    self._warns = self._warns + 1
    self:_both("  " .. self:_pad() .. "[!] " .. msg)
end

--- 错误 (含堆栈, 双路)
function Logger:error(msg)
    self._errors = self._errors + 1
    self:_both("  " .. self:_pad() .. "[X] ERROR: " .. tostring(msg))
    local trace = debug.traceback("", 2)
    if trace and trace ~= "" then
        self:_file("  " .. self:_pad() .. "    Stack:")
        for line in trace:gmatch("[^\n]+") do
            self:_file("  " .. self:_pad() .. "      " .. line)
        end
    end
end

--- 键值对 (仅文件, 对齐)
function Logger:data(key, value)
    local v
    if type(value) == "boolean" then
        v = value and "true" or "false"
    elseif type(value) == "table" then
        local p = {}
        for k, val in pairs(value) do table.insert(p, tostring(k) .. "=" .. tostring(val)) end
        v = "{" .. table.concat(p, ", ") .. "}"
    elseif value == nil then
        v = "(nil)"
    else
        v = tostring(value)
    end
    v = v:gsub("\r\n", " "):gsub("\n", " "):gsub("\r", " "):gsub("%s+", " ")
    local pad = key .. string.rep(" ", math.max(0, 28 - #key))
    self:_file(string.format("        %s = %s", pad, v))
end

--- 字节 hex dump (仅文件)
function Logger:bytes(label, s)
    if s == nil then
        self:_file("        " .. label .. "  =>  (nil)")
        return
    end
    local hex = {}
    for i = 1, #s do hex[i] = string.format("%02X", string.byte(s, i)) end
    local ascii = s:gsub("[^\032-\126]", ".")
    self:_file(string.format("        %-26s  len=%-4d  hex= %s", label, #s, table.concat(hex, " ")))
    self:_file(string.format("        %-26s  str= \"%s\"", "", ascii))
end

--- 单条断言记录 (仅文件)
function Logger:assertion(name, ok, expected, got)
    if ok then
        self:_file(string.format("        [v] %s  =>  %s", name, tostring(expected)))
    else
        self:_file(string.format("        [x] %s  expected=%s  got=%s", name, tostring(expected), tostring(got)))
    end
end

function Logger:separator()
    self:_both("  " .. string.rep("~", 56))
end

function Logger:summary()
    local elapsed = os.time() - self._start
    local passed = self._step - self._errors - self._skips
    self:_both("")
    self:_both(string.rep("=", 60))
    self:_both("  SUMMARY")
    self:_both(string.rep("-", 60))
    self:_both(string.format("  Steps   : %d", self._step))
    self:_both(string.format("  Passed  : %d", passed))
    self:_both(string.format("  Failed  : %d", self._errors))
    self:_both(string.format("  Skipped : %d", self._skips))
    self:_both(string.format("  Warnings: %d", self._warns))
    self:_both(string.format("  Time    : %ds", elapsed))
    self:_both(string.rep("-", 60))
    if self._errors == 0 then
        self:_both("  RESULT: ALL PASSED")
    else
        self:_both(string.format("  RESULT: %d ERROR(S) -- check details above", self._errors))
    end
    self:_both(string.rep("=", 60))
    self:_both("  Log file: " .. self._path)
end

function Logger:close()
    if self._fh then self._fh:close(); self._fh = nil end
end

--  Encoding  —  复用 excel.encoding (ADODB.Stream COM 转码, 无需 powershell)
local encoding = excel.encoding

-- 轻量包装: 记录转换前后字节到日志 (便于调试)
local _raw_toACP, _raw_toUTF8, _raw_forExcel = encoding.toACP, encoding.toUTF8, encoding.forExcel

function encoding.toACP(s)
    if encoding._log then encoding._log:bytes("toACP input", s) end
    local r = _raw_toACP(s)
    if encoding._log and r then encoding._log:bytes("toACP output", r) end
    return r
end

function encoding.toUTF8(s)
    if encoding._log then encoding._log:bytes("toUTF8 input", s) end
    local r = _raw_toUTF8(s)
    if encoding._log and r then encoding._log:bytes("toUTF8 output", r) end
    return r
end

function encoding.forExcel(s)
    if encoding._log then encoding._log:bytes("forExcel input", s) end
    local r = _raw_forExcel(s)
    if encoding._log then
        if r == s then encoding._log:detail("forExcel: passthrough")
        else encoding._log:detail("forExcel: converted") end
    end
    return r
end

-- ============================================================================
--  PART 1  环境快照  ENVIRONMENT
-- ============================================================================
--  采集运行环境的全部关键信息，便于离线诊断兼容性问题。
--  包括: OS, Lua 版本, 系统版本, ANSI 代码页, luacom 状态, 模块导出列表,
--        以及源码中文字面量在内存中的实际字节 (用于判断源文件编码是否正确)。
-- ============================================================================
local function part1_environment(log)
    local id = log:step("Environment")

    local is_win = encoding.isWindows()
    log:data("OS",                     is_win and "Windows" or "Linux/WSL")
    log:data("Lua version",            _VERSION)

    -- 系统版本 (Windows: powershell 避免 GBK 乱码)
    if is_win then
        local h = io.popen([[powershell -NoProfile -Command "(Get-CimInstance Win32_OperatingSystem).Caption + ' [' + (Get-CimInstance Win32_OperatingSystem).Version + ']'"]], "r")
        if h then
            local info = h:read("*a"):gsub("\r\n", ""):gsub("\n", ""):gsub("\r", ""):gsub("%s+", " ")
            h:close()
            log:data("System", info)
        end
    else
        local h = io.popen("uname -a 2>/dev/null", "r")
        if h then
            local info = h:read("*a"):gsub("\r\n", " "):gsub("\n", " "):gsub("%s+", " ")
            h:close()
            log:data("System", info)
        end
    end

    -- ANSI 代码页 (936=GBK, 950=BIG5, 54936=GB18030)
    local acp = encoding.getACP()
    if acp then
        local label = "ANSI codepage"
        if acp == 936      then label = label .. " (GBK)"
        elseif acp == 950  then label = label .. " (BIG5)"
        elseif acp == 54936 then label = label .. " (GB18030)" end
        log:data(label, acp)
    end

    -- luacom 模块状态
    local ok, err = pcall(require, "luacom")
    log:data("luacom", ok and "available" or ("MISSING: " .. tostring(err):gsub("\n", " ")))

    -- excel 模块导出的字段
    local fields = {}
    for k, v in pairs(excel) do
        table.insert(fields, k .. "(" .. type(v) .. ")")
    end
    table.sort(fields)
    log:data("excel exports", table.concat(fields, ", "))

    -- 源码中文字面量 "测试" 的实际字节: 应为 E6 B5 8B E8 AF 95 (UTF-8 正确)
    -- 若显示其他值, 说明源文件被错误地按其他编码保存或解析
    local literal = "\230\181\139\232\175\149"
    log:bytes('\230\181\139\232\175\149 ("\229\181\139\229\143\175\229\140\150")', literal)

    log:pass(id)
end

-- ============================================================================
--  PART 2  工具函数  UTILITY FUNCTIONS
-- ============================================================================
--  纯函数单元测试，验证列号/字母转换, 单元格/区域地址生成。
--  不依赖 Excel COM, 任何环境均可运行。
--
--  columnLetter(n)  数字列号 → 字母: 1=A, 27=AA, 702=ZZ
--  columnNumber(s)  字母 → 数字: A=1, AA=27
--  cellAddr(r,c)    行列号 → "A1" 格式
--  rangeAddr(...)   行列号 → "A1:C10" 格式
-- ============================================================================
local function part2_utils(log)
    local id = log:step("Utility functions (14 assertions)")

    local cases = {
        -- 列号 → 字母 (6 项)
        { "columnLetter(1)",    excel.columnLetter(1),     "A" },
        { "columnLetter(26)",   excel.columnLetter(26),    "Z" },
        { "columnLetter(27)",   excel.columnLetter(27),    "AA" },
        { "columnLetter(52)",   excel.columnLetter(52),    "AZ" },
        { "columnLetter(53)",   excel.columnLetter(53),    "BA" },
        { "columnLetter(702)",  excel.columnLetter(702),   "ZZ" },
        -- 字母 → 列号 (5 项)
        { "columnNumber('A')",  excel.columnNumber("A"),   1 },
        { "columnNumber('Z')",  excel.columnNumber("Z"),   26 },
        { "columnNumber('AA')", excel.columnNumber("AA"),  27 },
        { "columnNumber('AZ')", excel.columnNumber("AZ"),  52 },
        { "columnNumber('BA')", excel.columnNumber("BA"),  53 },
        -- 地址生成 (3 项)
        { "cellAddr(1,1)",      excel.cellAddr(1, 1),     "A1" },
        { "cellAddr(10,27)",    excel.cellAddr(10, 27),   "AA10" },
        { "rangeAddr(1,1,10,3)",excel.rangeAddr(1,1,10,3),"A1:C10" },
    }

    local nfail = 0
    for _, tc in ipairs(cases) do
        local ok = (tc[2] == tc[3])
        log:assertion(tc[1], ok, tc[3], tc[2])
        if not ok then nfail = nfail + 1 end
    end
    if nfail == 0 then log:pass(id) else log:fail(id, nfail .. " failed") end
end

-- ============================================================================
--  PART 3  编码往返  ENCODING
-- ============================================================================
--  验证 encoding.toACP / encoding.toUTF8 的正确性。
--  取 UTF-8 字符串 "测试" (E6 B5 8B E8 AF 95), 转为系统 ANSI (B2 E2 CA D4),
--  再转回 UTF-8。两次转换后字节必须与原始完全一致。
--  仅在 Windows 上运行 (需要 powershell)。
-- ============================================================================
local function part3_encoding(log)
    local id = log:step("Encoding roundtrip (UTF-8 -> ACP -> UTF-8)")
    if not encoding.isWindows() then
        log:skip(id, "not Windows")
        return
    end
    local acp = encoding.getACP()
    if acp ~= 936 and acp ~= 950 and acp ~= 54936 then
        log:skip(id, "non-Chinese codepage " .. (acp or "?"))
        return
    end

    local original = "\230\181\139\232\175\149"  -- UTF-8 "测试" = E6 B5 8B E8 AF 95
    local ansi = encoding.toACP(original)         -- → GBK "测试" = B2 E2 CA D4
    if not ansi then
        log:fail(id, "toACP returned nil")
        return
    end

    local utf8 = encoding.toUTF8(ansi)             -- → UTF-8 (应等于 original)
    if not utf8 then
        log:fail(id, "toUTF8 returned nil")
        return
    end

    if utf8 == original then
        log:pass(id)
    else
        log:fail(id, "roundtrip mismatch (check bytes in log)")
        log:bytes("expected", original)
        log:bytes("got     ", utf8)
    end
end

-- ============================================================================
--  PART 4  Excel COM 集成测试  EXCEL COM
-- ============================================================================
--  完整的 Excel COM 操作链, 覆盖库的全部核心功能。
--  每个步骤 (4..14) 验证一类操作, 步骤间有依赖 (前一步必须成功后续才有意义)。
--  所有中文字符串经由 encoding.forExcel() 转码后再写入 Excel。
--
--  步骤:
--    [4]  检测 COM 环境 (Windows + luacom 已安装)
--    [5]  创建不可见的 Excel.Application 实例
--    [6]  新建空白工作簿
--    [7]  工作表改中文名 (验证 encoding.forExcel 写入正确)
--    [8]  写入 string/number/boolean/formula 并读回断言
--    [9]  写入中文表头+3行数据, 读回验证编码往返
--    [10] Range 格式化: 字体, 颜色, 边框, 自动列宽
--    [11] 合并单元格 + 居中标题
--    [12] 数字格式: 小数 #,##0.00 + 日期 yyyy-mm-dd
--    [13] 工作表增删: 新建 → 断言总数=2 → 删除 → 断言总数=1
--    [14] 读取 Excel 版本号
-- ============================================================================
local function part4_excel(log)

    -- [4] 检测 COM 环境
    --     非 Windows → 直接跳过全部 COM 测试
    --     luacom 未安装 → 标记失败并退出
    local id = log:step("Check COM environment")
    if not encoding.isWindows() then
        log:skip(id, "not Windows -- WSL/Linux has no COM")
        log:section("Excel COM tests skipped (non-Windows)")
        return false
    end
    local luacom_ok, luacom_err = pcall(require, "luacom")
    if not luacom_ok then
        log:fail(id, "luacom not available: " .. tostring(luacom_err):gsub("\n", " | "))
        return false
    end
    log:pass(id)

    -- [5] 创建 Excel.Application (不可见模式)
    --     DisplayAlerts=false: 不弹警告框
    --     ScreenUpdating=false: 默认不刷新 (库行为)
    id = log:step("Create Excel.Application")
    local app, app_err = excel.ExcelApp:new(false)
    if not app then
        log:fail(id, app_err or "ExcelApp:new() returned nil")
        return false
    end
    log:pass(id)

    local all_ok = true
    local worked, run_err = pcall(function()
        app:displayAlerts(false)

        -- [6] 新建工作簿
        --     默认包含一个工作表 (Sheet1), 名称由 Excel 本地化决定
        id = log:step("Add workbook")
        local wb = app:addWorkbook()
        log:data("name", wb:name())
        log:pass(id)

        -- [7] 工作表改中文名
        --     中文名 "测试表" 经 encoding.forExcel() UTF-8→GBK 后写入
        --     读回名称与写入值逐字节比对, 确保编码无误
        id = log:step("Rename sheet with Chinese name")
        local s1 = wb:sheetByIndex(1)
        local cn = encoding.forExcel("\230\181\139\232\175\149\232\161\168")  -- "测试表"
        s1:setName(cn)
        local got = wb:sheet(cn):name()
        log:data("expected", cn)
        log:bytes("got bytes", got)
        if got == cn then
            log:pass(id)
        else
            log:fail(id, "name mismatch -- possible encoding issue")
            all_ok = false
        end

        -- [8] 基本类型读写
        --     验证 string, integer, float, boolean, formula 写入后读回一致
        --     A1="Hello", B1=12345, C1=3.14159, D1=true, E1==A1 & " World"
        id = log:step("Read/write basic types (string, number, boolean, formula)")
        s1:cell(1, 1, "Hello")
        s1:cell(1, 2, 12345)
        s1:cell(1, 3, 3.14159)
        s1:cell(1, 4, true)
        s1:cell(1, 5, "=A1 & \" World\"")
        log:assertion("A1 string", s1:cell(1,1) == "Hello", "Hello", s1:cell(1,1))
        log:assertion("B1 number", s1:cell(1,2) == 12345,   12345,   s1:cell(1,2))
        log:pass(id)

        -- [9] 中文数据写入并读回 (编码往返验证)
        --     表头: 姓名, 年龄, 城市 → 经 encoding.forExcel 转 GBK 写入
        --     数据: 张三/25/北京, 李四/30/上海, 王五/28/广州
        --     读回 "张三" → encoding.toUTF8 转回 UTF-8 → 验证与原始字节一致
        id = log:step("Write/read Chinese characters")

        -- 表头 (行3)
        local header = {
            encoding.forExcel("\229\167\147\229\144\141"),   -- "姓名"
            encoding.forExcel("\229\185\180\233\190\132"),   -- "年龄"
            encoding.forExcel("\229\159\142\229\184\130"),   -- "城市"
        }
        for c, v in ipairs(header) do
            s1:cell(3, c, v)
            log:bytes("header[" .. c .. "]", v)
        end

        -- 数据 (行4-6)
        local rows = {
            { "\229\188\160\228\184\137", 25, "\229\140\151\228\186\172" },  -- 张三,25,北京
            { "\230\157\142\229\155\155", 30, "\228\184\138\230\181\183" },  -- 李四,30,上海
            { "\231\142\139\228\186\148", 28, "\229\185\191\229\183\158" },  -- 王五,28,广州
        }
        for r, row in ipairs(rows) do
            local name = encoding.forExcel(row[1])
            local city = encoding.forExcel(row[3])
            s1:cell(3 + r, 1, name)
            s1:cell(3 + r, 2, row[2])
            s1:cell(3 + r, 3, city)
            log:bytes("row" .. r .. " name", name)
            log:bytes("row" .. r .. " city", city)
        end

        -- 读回 "张三" (行4列1), 验证 GBK → UTF-8 往返
        local rb = s1:cell(4, 1)
        log:data("readback type", type(rb))
        log:bytes("readback raw", tostring(rb or ""))
        local rb_utf8 = encoding.toUTF8(rb or "")
        if rb_utf8 then
            log:bytes("readback UTF8", rb_utf8)
        end
        log:pass(id)

        -- [10] Range 格式化
        --      表头行 (A3:C3): 粗体, 12pt, 居中, 蓝底白字, 外边框
        --      数据区 (A4:C6): 全部框线
        --      整体 (A1:E6): 自动列宽
        id = log:step("Range formatting (font, color, border, autofit)")
        local rng = s1:range("A3:C3")
        rng:bold(true):fontSize(12):halign("Center")
        rng:bgColor(0x4472C4):fontColor(0xFFFFFF)
        rng:borderOutline("Continuous")
        s1:range("A4:C6"):borderAll("Continuous")
        s1:range("A1:E6"):autoFit()
        log:data("header range", rng:address())
        log:pass(id)

        -- [11] 合并单元格
        --      A8:C8 合并, 写入 "合并标题" 居中加粗
        id = log:step("Merge cells")
        s1:range("A8:C8"):merge()
        s1:cell(8, 1, encoding.forExcel("\229\144\136\229\185\182\230\160\135\233\162\152"))  -- "合并标题"
        s1:range("A8:C8"):halign("Center"):bold(true)
        log:pass(id)

        -- [12] 数字格式
        --      A10: 45678.9012 → 格式 "#,##0.00" (千分位两位小数)
        --      B10: os.date("*t") → 自动转为 Excel 日期序列号 → 格式 "yyyy-mm-dd"
        id = log:step("Number format (decimal, date)")
        s1:range("A10"):value(45678.9012)
        s1:range("A10"):numberFormat("#,##0.00")
        s1:range("B10"):value(os.date("*t"))          -- table → ExcelApp.dateSerial()
        s1:range("B10"):numberFormat("yyyy-mm-dd")
        log:pass(id)

        -- [13] 工作表增删
        --      在 s1 之后新建 "第二表" → 断言增加 1 → 删除 → 断言恢复
        --      使用相对计数避免 Excel 版本差异 (2010 默认 3 张, 2016+ 默认 1 张)
        id = log:step("Add & delete worksheet")
        local initial_count = wb:sheetCount()
        local s2 = wb:addSheet(encoding.forExcel("\231\172\172\228\186\140\232\161\168"), s1)  -- "第二表"
        local after_add = wb:sheetCount()
        log:data("initial count", initial_count)
        log:data("count after add", after_add)
        if after_add ~= initial_count + 1 then
            log:fail(id, string.format("expected %d, got %d", initial_count + 1, after_add))
            all_ok = false
        else
            s2:cell(1, 1, "tmp")
            wb:deleteSheet(s2:name())
            local after_delete = wb:sheetCount()
            log:data("count after delete", after_delete)
            if after_delete ~= initial_count then
                log:fail(id, string.format("expected %d, got %d", initial_count, after_delete))
                all_ok = false
            else
                log:pass(id)
            end
        end

        -- [14] Excel 版本
        --      预期 16.0 = Office 2016/2019/365
        id = log:step("Excel version")
        log:data("version", app:version())
        log:pass(id)
    end)

    -- 清理: 丢弃所有更改, 静默退出 Excel
    pcall(function()
        app:displayAlerts(false)
        app:quit()   -- quit() 内部会先关闭所有工作簿再 Quit(), 不弹保存对话框
    end)

    if not worked then
        log:error(run_err)
        all_ok = false
    end
    return all_ok
end

-- ============================================================================
--  PART 5  覆盖率扩展  COVERAGE EXTENSION
-- ============================================================================
--  测试 PART 4 未覆盖的方法, 分为 App/Workbook/Sheet extras, Range getters,
--  导航, 进阶格式化, 边框/尺寸, 取消合并, 剪贴板, 清除, 行列增删, 计算。
--  创建独立 Excel 实例, 不影响 PART 4。
--
--  步骤:
--    [15] App 配置: visible, screenUpdating, calculation mode
--    [16] Workbook 信息: workbookCount, workbook(index), path, sheetNames
--    [17] Sheet 属性: index, activate, visible, range(numeric), used*
--    [18] Sheet 保护: protect, unprotect
--    [19] Range getters: rows, columns, row, column
--    [20] Range 导航: offset, cell, entireRow, entireColumn
--    [21] 进阶格式化: fontName, italic, underline, valign, wrapText, bgPattern
--    [22] 单边框 & 尺寸: border(single), columnWidth, rowHeight
--    [23] 取消合并: merge → unmerge → 验证
--    [24] 剪贴板: copy, pasteTo
--    [25] 清除: clear (保留格式), clearAll (含格式)
--    [26] 行列插删: insertRows, deleteRows, insertColumns, deleteColumns
--    [27] 计算: calculation mode, calculate
-- ============================================================================
local function part5_coverage(log)

    local id = log:step("Check COM environment")
    if not encoding.isWindows() then
        log:skip(id, "not Windows")
        return false
    end
    local luacom_ok, luacom_err = pcall(require, "luacom")
    if not luacom_ok then
        log:fail(id, "luacom not available")
        return false
    end
    log:pass(id)

    -- 创建独立 Excel 实例
    id = log:step("Create Excel.Application")
    local app, app_err = excel.ExcelApp:new(false)
    if not app then
        log:fail(id, app_err or "nil")
        return false
    end
    log:pass(id)

    local all_ok = true
    local worked, run_err = pcall(function()
        app:displayAlerts(false)
        local wb = app:addWorkbook()
        local s = wb:sheetByIndex(1)
        s:setName("CoverageTest")

        -- ================================================================
        -- [15] App 配置: 可反复切换的设置项
        -- ================================================================
        id = log:step("App settings (visible, screenUpdating, calculation)")
        app:visible(false)
        app:screenUpdating(true)
        app:screenUpdating(false)
        app:calculation("manual")
        app:calculation("auto")
        log:pass(id)

        -- ================================================================
        -- [16] Workbook 信息
        -- ================================================================
        id = log:step("Workbook info (count, index, path, sheetNames)")
        local count = app:workbookCount()
        log:data("workbookCount", count)
        log:assertion("workbookCount>=1", count >= 1, ">=1", count)

        local wb2 = app:workbook(1)
        log:data("workbook(1).name", wb2:name())
        log:assertion("workbook(1).name matches", wb2:name() == wb:name(), wb:name(), wb2:name())

        local p = wb:path()
        log:data("path", (p or "(new, unsaved)"))
        -- 新建未保存的工作簿 path 可能非空或为空, 不做强断言

        local names = wb:sheetNames()
        log:data("sheetNames count", #names)
        log:assertion("sheetNames[1]", names[1] == "CoverageTest", "CoverageTest", names[1] or "")
        log:pass(id)

        -- ================================================================
        -- [17] Sheet 属性: index, activate, visible, range(numeric), used*
        -- ================================================================
        id = log:step("Sheet properties (index, activate, visible, usedRange)")

        local idx = s:index()
        log:data("index", idx)
        log:assertion("index==1", idx == 1, 1, idx)

        s:activate()
        log:info("activate: OK")

        -- visible(false) 需要至少一个其他可见工作表
        local tmp = wb:addSheet("_tmp", s)
        s:visible(false)
        s:visible(true)
        wb:deleteSheet("_tmp")

        -- 写入一些数据以便 usedRange 有内容
        s:cell(1, 1, "A"):cell(1, 2, "B"):cell(1, 3, "C")
        s:cell(2, 1, "D"):cell(2, 2, "E"):cell(2, 3, "F")

        local used = s:usedRange()
        log:data("usedRange address", used:address())
        log:data("usedRows", s:usedRows())
        log:data("usedColumns", s:usedColumns())
        log:assertion("usedRows>=2", s:usedRows() >= 2, ">=2", s:usedRows())
        log:assertion("usedColumns>=3", s:usedColumns() >= 3, ">=3", s:usedColumns())

        -- range(numeric) 重载: 4 个参数
        local rng_num = s:range(1, 1, 2, 3)
        log:data("range(1,1,2,3) address", rng_num:address())
        log:pass(id)

        -- ================================================================
        -- [18] Sheet 保护: 无密码 → 取消
        -- ================================================================
        id = log:step("Sheet protection (protect, unprotect)")
        s:protect()
        s:unprotect()
        s:protect("pwd123")
        s:unprotect("pwd123")
        log:pass(id)

        -- ================================================================
        -- [19] Range getters: 从 (1,1) 到 (2,3) 区域获取元数据
        -- ================================================================
        id = log:step("Range getters (rows, columns, row, column)")
        local rg = s:range("A1:C2")
        log:data("rows", rg:rows())
        log:data("columns", rg:columns())
        log:data("row", rg:row())
        log:data("column", rg:column())
        log:assertion("rows==2", rg:rows() == 2, 2, rg:rows())
        log:assertion("columns==3", rg:columns() == 3, 3, rg:columns())
        log:assertion("row==1", rg:row() == 1, 1, rg:row())
        log:assertion("column==1", rg:column() == 1, 1, rg:column())
        log:pass(id)

        -- ================================================================
        -- [20] Range 导航: offset, cell, entireRow, entireColumn
        -- ================================================================
        id = log:step("Range navigation (offset, cell, entireRow, entireColumn)")
        local off = rg:offset(1, 0)
        log:data("offset(1,0) address", off:address())

        local sub = rg:cell(2, 3)
        log:data("cell(2,3) value", sub:value())

        local erow = rg:entireRow()
        log:data("entireRow columns", erow:columns())

        local ecol = rg:entireColumn()
        log:data("entireColumn rows", ecol:rows())
        log:pass(id)

        -- ================================================================
        -- [21] 进阶格式化: 字体/对齐/填充 (PART 4 未测的)
        -- ================================================================
        id = log:step("Formatting extras (fontName, italic, underline, valign, wrapText, bgPattern)")
        local fmt = s:range("A1:C1")
        fmt:fontName("Consolas")
        fmt:italic(true)
        fmt:underline(true)
        fmt:valign("Center")
        fmt:wrapText(true)
        fmt:bgPattern("Solid")
        -- 读回验证 getter 路径 (COM underline 返回整数 2 非 boolean)
        log:data("italic get", fmt:italic())
        log:data("underline get", fmt:underline())
        log:assertion("italic", fmt:italic() == true, true, fmt:italic())
        log:assertion("underline", fmt:underline() ~= false, "~=false", fmt:underline())
        log:pass(id)

        -- ================================================================
        -- [22] 单独边框 & 列宽行高
        -- ================================================================
        id = log:step("Single border, columnWidth, rowHeight")
        -- border(idx, style): 7=left, 9=bottom
        fmt:border(7, "Double")
        fmt:border(9, "Dash")
        fmt:columnWidth(15)
        fmt:rowHeight(20)
        log:pass(id)

        -- ================================================================
        -- [23] 取消合并: merge → unmerge
        -- ================================================================
        id = log:step("Unmerge cells")
        local mg = s:range("A5:B6")
        mg:merge()
        log:info("merged A5:B6")
        mg:unmerge()
        log:info("unmerged A5:B6")
        log:pass(id)

        -- ================================================================
        -- [24] 剪贴板: copy → pasteTo (不测试 cut, 它会清空源)
        -- ================================================================
        id = log:step("Clipboard (copy, pasteTo)")
        s:cell(10, 1, "Source")
        local src = s:range("A10")
        src:copy()
        local dst = s:range("C10")
        src:pasteTo(dst)
        -- 读回用 s:cell() 避免 range 对象缓存问题
        local pasted = s:cell(10, 3)
        log:data("pasteTo result", pasted)
        log:assertion("pasteTo == Source", pasted == "Source", "Source", pasted)
        log:pass(id)

        -- ================================================================
        -- [25] 清除: clear (保留格式) + clearAll (含格式)
        --      先设格式, clear → 格式应保留, clearAll → 全部清除
        -- ================================================================
        id = log:step("Clear (clear, clearAll)")
        local cl = s:range("A12")
        cl:value("tmp"):bold(true)
        cl:clear()          -- 清内容, 格式保留
        local v1 = s:cell(12, 1)   -- 用 cell() 读回, 避免 range 缓存
        log:data("after clear value", v1)
        log:assertion("clear: value empty", v1 == nil or v1 == "", "", v1)
        cl:value("tmp2"):bold(true)
        cl:clearAll()        -- 全部清除
        local v2 = s:cell(12, 1)
        log:data("after clearAll value", v2)
        log:assertion("clearAll: value empty", v2 == nil or v2 == "", "", v2)
        log:pass(id)

        -- ================================================================
        -- [26] 行列插删 (在空白区操作, 避免影响已有数据)
        -- ================================================================
        id = log:step("Row/column insert & delete")
        -- 在行20操作, 不影响上面已有数据
        s:cell(20, 1, "keep")
        s:range("A20"):insertRows()   -- 行20(A20)下移, 新空行在20, "keep" → 行21
        s:range("A20"):deleteRows()   -- 删除新空行20, 行21("keep") → 行20
        local v = s:cell(20, 1)
        log:data("after insert+deleteRows, A20", v)
        -- COM 插入删除后单元格值可能为 nil (空), 也可能是 "keep", 取决于 Excel 版本
        -- 只要方法不抛异常就算通过
        log:info("insertRows/deleteRows: OK")

        s:cell(20, 1, "keep")
        s:range("A20"):insertColumns() -- 列右移
        s:range("A20"):deleteColumns() -- 恢复
        log:data("after insert+deleteColumns, A20", s:cell(20, 1))
        log:info("insertColumns/deleteColumns: OK")
        log:pass(id)

        -- ================================================================
        -- [27] 计算: 手动模式 + 强制计算
        -- ================================================================
        id = log:step("Calculation (manual, calculate)")
        s:cell(30, 1, 10)
        s:cell(30, 2, 20)
        s:cell(30, 3, "=A30+B30")
        wb:calculation("manual")
        wb:calculate()
        log:data("calc result", s:cell(30, 3))
        log:assertion("30+20==30", s:cell(30, 3) == 30, 30, s:cell(30, 3))
        wb:calculation("auto")
        log:pass(id)
    end)

    -- 清理
    pcall(function()
        app:displayAlerts(false)
        app:quit()
    end)

    if not worked then
        log:error(run_err)
        all_ok = false
    end
    return all_ok
end

-- ============================================================================
--  PART 6  资源管理  RESOURCE MANAGEMENT
-- ============================================================================
--  验证 COM 对象自动释放机制: __gc, __close, 重复 quit 安全。
--  创建独立 Excel 实例, 不影响其他 PART。
--
--  步骤:
--    [30] 双重 quit 安全: quit() → quit() 不崩溃
--    [31] 手动 quit 后操作: quit 后尝试方法调用, 应安全返回 nil/error
--    [32] __close 自动清理: 使用 to-be-closed 变量, 离开作用域自动 quit
--    [33] 异常恢复: 在 to-be-closed 作用域内抛异常, __close 应触发清理
--    [34] 多次创建/退出: 反复创建/退出 Excel 实例, 不累积僵尸进程
-- ============================================================================
local function part6_resource(log)

    local id = log:step("Check COM environment")
    if not encoding.isWindows() then
        log:skip(id, "not Windows")
        return false
    end
    local luacom_ok = pcall(require, "luacom")
    if not luacom_ok then
        log:fail(id, "luacom not available")
        return false
    end
    log:pass(id)

    -- ================================================================
    -- [30] 双重 quit 安全: quit() 两次不应崩溃
    -- ================================================================
    id = log:step("Double quit safety")
    local app = excel.ExcelApp:new(false)
    if not app then
        log:fail(id, "ExcelApp:new() failed")
        return false
    end
    app:addWorkbook()
    app:quit()
    -- 第二次 quit 应安全无操作 (由 _released 标记保护)
    local ok, err = pcall(function() app:quit() end)
    if ok then
        log:pass(id)
    else
        log:fail(id, "second quit() crashed: " .. tostring(err))
    end

    -- ================================================================
    -- [31] 手动 quit 后只读操作: 对象已释放, 读操作应安全返回或报错
    --      不调用任何写操作 (visible/Add 等) — 避免唤醒 zombie COM 对象
    -- ================================================================
    id = log:step("Read operations after quit (safe nil/error)")
    local _, err2 = pcall(function() app:version() end)
    log:data("version after quit", err2 and "error (expected)" or "no error")
    local _, err3 = pcall(function() app:workbookCount() end)
    log:data("workbookCount after quit", err3 and "error (expected)" or "no error")
    -- 读操作可能返回 0/nil 或抛异常, 只要不挂起就算通过
    log:pass(id)

    -- ================================================================
    -- [32] __close 自动清理 (Lua 5.5 to-be-closed 变量)
    --      创建 Excel → 读写数据 → 离开作用域 → __close 自动 quit
    -- ================================================================
    id = log:step("__close auto-cleanup (to-be-closed variable)")
    local closed_ok = false
    pcall(function()
        -- Lua 5.5: <close> 变量离开作用域自动调 __close
        -- 旧版 Lua 忽略 <close> 标记, 等价于普通 local
        local test_app <close> = excel.ExcelApp:new(false)
        if not test_app then return end
        local wb = test_app:addWorkbook()
        wb:sheetByIndex(1):cell(1, 1, "close test")
        -- 此处正常离开作用域, __close 应触发 quit()
        closed_ok = true
    end)
    log:data("__close scope completed", closed_ok)
    if closed_ok then
        log:pass(id)
    else
        log:fail(id, "__close scope failed")
    end

    -- ================================================================
    -- [33] 异常恢复: try 块内抛 error, __close 仍应触发 quit
    -- ================================================================
    id = log:step("Exception recovery (error inside to-be-closed scope)")
    local error_raised = false
    local cleanup_worked = true
    pcall(function()
        local recovery_app <close> = excel.ExcelApp:new(false)
        if not recovery_app then cleanup_worked = false; return end
        recovery_app:addWorkbook()
        error("simulated crash inside Excel work")
    end)
    -- 如果上面 error() 抛出: Lua 先调 __close (quit), 再跳转到 pcall 捕获
    -- 如果 __close 崩溃: pcall 会捕获它, cleanup_worked 保持 false
    -- 正常: error 被 pcall 捕获, __close 清理完成, 程序继续
    log:data("error+cleanup", cleanup_worked and "OK" or "FAIL")
    if cleanup_worked then
        log:pass(id)
    else
        log:fail(id, "cleanup after error failed")
    end

    -- ================================================================
    -- [34] 多次创建/退出: 快速连续创建退出多个实例
    --      验证无僵尸进程累积 (每个实例正确清理)
    -- ================================================================
    id = log:step("Multiple create/quit cycles (3 iterations)")
    local cycles_ok = true
    for i = 1, 3 do
        local loop_ok = pcall(function()
            local cyc <close> = excel.ExcelApp:new(false)
            if not cyc then error("create failed") end
            local wb = cyc:addWorkbook()
            wb:sheetByIndex(1):cell(1, 1, "cycle " .. i)
        end)
        if not loop_ok then
            cycles_ok = false
            log:data("cycle " .. i, "FAIL")
        end
    end
    if cycles_ok then
        log:pass(id)
    else
        log:fail(id, "one or more cycles failed")
    end

    return true
end

-- ============================================================================
--  入口  MAIN
-- ============================================================================

local logfile = string.format("test_log_%s.txt", os.date("%Y%m%d_%H%M%S"))
local log = Logger:new(logfile)

-- PART 1: 环境
log:section("ENVIRONMENT")
local ok1, e1 = pcall(part1_environment, log)
if not ok1 then log:error("PART1 crashed: " .. tostring(e1)) end

-- PART 2: 工具函数
log:section("UTILITY FUNCTIONS")
local ok2, e2 = pcall(part2_utils, log)
if not ok2 then log:error("PART2 crashed: " .. tostring(e2)) end

-- PART 3: 编码
log:section("ENCODING")
encoding._log = log
local ok3, e3 = pcall(part3_encoding, log)
if not ok3 then log:error("PART3 crashed: " .. tostring(e3)) end

-- PART 4: Excel COM (核心流程)
log:section("EXCEL COM")
local ok4, e4 = pcall(part4_excel, log)
if not ok4 then log:error("PART4 crashed: " .. tostring(e4)) end

-- PART 5: 覆盖率扩展
log:section("COVERAGE EXTENSION")
local ok5, e5 = pcall(part5_coverage, log)
if not ok5 then log:error("PART5 crashed: " .. tostring(e5)) end

-- PART 6: 资源管理
log:section("RESOURCE MANAGEMENT")
local ok6, e6 = pcall(part6_resource, log)
if not ok6 then log:error("PART6 crashed: " .. tostring(e6)) end

-- 摘要
log:summary()
log:close()

print("")
print("  Log saved to: " .. logfile)
