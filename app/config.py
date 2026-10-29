import os
import tempfile
from pathlib import Path

# 基础配置
BASE_DIR = Path(__file__).resolve().parent.parent

# Windows / Linux 都可用的临时目录
_default_upload = Path(tempfile.gettempdir()) / "bank_uploads"
UPLOAD_DIR = Path(os.getenv("UPLOAD_DIR", str(_default_upload)))
#UPLOAD_DIR = Path(os.getenv("UPLOAD_DIR", "/tmp/bank_uploads"))
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)

MAX_FILE_SIZE_MB = int(os.getenv("MAX_FILE_SIZE_MB", "20"))
MAX_FILE_SIZE = MAX_FILE_SIZE_MB * 1024 * 1024
MAX_FILES = int(os.getenv("MAX_FILES", "5"))

ALLOWED_EXTENSIONS = {".pdf", ".xlsx", ".xls", ".csv"}

# 隐私：处理后立即删除，不留存任何用户文件超过 10 分钟
FILE_TTL_SECONDS = 600

# 支持的银行关键词（用于自动识别）
BANK_KEYWORDS = {
    "工商银行": ["工商银行", "ICBC", "中国工商银行"],
    "建设银行": ["建设银行", "CCB", "中国建设银行"],
    "农业银行": ["农业银行", "ABC", "中国农业银行"],
    "中国银行": ["中国银行", "BOC"],
    "招商银行": ["招商银行", "CMB", "招商"],
    "交通银行": ["交通银行", "BOCOM", "交行"],
    "中信银行": ["中信银行", "CITIC"],
    "浦发银行": ["浦发银行", "SPDB", "浦东发展"],
    "民生银行": ["民生银行", "CMBC"],
    "兴业银行": ["兴业银行", "CIB"],
    "平安银行": ["平安银行", "PAB"],
    "光大银行": ["光大银行", "CEB"],
    "广发银行": ["广发银行", "CGB"],
    "邮储银行": ["邮储银行", "PSBC", "邮政储蓄"],
    "通用": [],
}

# 交易类型关键词（简单分类）
INCOME_KEYWORDS = ["入账", "转入", "收入", "到账", "工资", "奖金", "退款", "利息", "收款", "贷"]
EXPENSE_KEYWORDS = ["支出", "转出", "消费", "支付", "扣款", "手续费", "还款", "借"]