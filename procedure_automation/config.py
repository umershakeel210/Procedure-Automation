# =========================
# Website URLs
# =========================


# =========================
# Procedure Code Settings
# =========================# =========================
# Login Settings
# =========================

LOGIN_URL = "https://rmi.vicenna.com/en/#/login"

LOGIN_USERNAME = "umar.shakeel@rmi.edu.pk"
LOGIN_PASSWORD = "umer@1122"

AUTO_LOGIN = True
AUTO_OPEN_CREATE_PROCEDURE = True

# Agar exact Create Procedure URL work karta ho to yahan paste karein.
# Agar direct URL work na kare to empty rehne dein. Code menu click se open karega.
CREATE_PROCEDURE_URL = "https://rmi.vicenna.com/en/#/manageProcedures/Add"


# =========================
# Login Selectors
# =========================

LOGIN_USERNAME_INPUT = "input[type='email'], input[name='email'], input[name='username'], input[ng-model*='user'], input[ng-model*='email']"
LOGIN_PASSWORD_INPUT = "input[type='password'], input[name='password'], input[ng-model*='password']"

LOGIN_BUTTON_XPATHS = [
    "//button[contains(normalize-space(), 'Login')]",
    "//button[contains(normalize-space(), 'Sign In')]",
    "//button[contains(normalize-space(), 'LOGIN')]",
    "//button[@type='submit']",
    "//*[contains(normalize-space(), 'Login') and (self::button or self::a or self::div)]"
]


# =========================
# Existing Form Selectors
# =========================

PROCEDURE_CODE_INPUT = "input[ng-model='procedure.Code']"
PROCEDURE_NAME_INPUT = "input[name='procedureName']"

DEFINED_AS_DROPDOWN = "md-select[name='definedas']"
DEPARTMENT_DROPDOWN = "md-select[name='category']"
SUB_DEPARTMENT_DROPDOWN = "md-select[name='subDepartment']"
RATE_LIST_DROPDOWN = "md-select[name='ratelist']"
PROCEDURE_TYPE_DROPDOWN = "md-select[name='procedureType']"

LENGTH_OF_STAY_INPUT = "input[ng-model='procedure.lengthOfStay']"
DESCRIPTION_INPUT = "input[ng-model='procedure.Description']"

PRICE_INPUT = "input[name='price']"
NOTES_INPUT = "textarea[name='notes']"

ACTIVE_CHECKBOX = "input[name='active']"

SAVE_BUTTON = "button[type='submit']"

SUCCESS_MESSAGE = ".toast-success"
ERROR_MESSAGE = ".toast-error"

START_PROCEDURE_CODE = 2257
PROCEDURE_CODE_PREFIX = "MSP-"

# =========================
# File Paths
# =========================

EXCEL_FILE_PATH = "data/procedures.xlsx"
LOG_FILE_PATH = "logs/automation_log.xlsx"


# =========================
# Main Form Selectors
# =========================
# Procedure Code ka selector screenshot se confirmed hai:
# <input name="procedureCode" ...>
PROCEDURE_CODE_INPUT = "input[ng-model='procedure.Code']"
PROCEDURE_NAME_INPUT = "input[name='procedureName']"
DEFINED_AS_DROPDOWN = "md-select[name='definedas']"
DEPARTMENT_DROPDOWN = "md-select[name='category']"
SUB_DEPARTMENT_DROPDOWN = "md-select[name='subDepartment']"
RATE_LIST_DROPDOWN = "md-select[name='ratelist']"
PROCEDURE_TYPE_DROPDOWN = "md-select[name='procedureType']"
LENGTH_OF_STAY_INPUT = "input[ng-model='procedure.lengthOfStay']"
DESCRIPTION_INPUT = "input[ng-model='procedure.Description']"

# ====
PRICE_INPUT = "input[name='price']"
NOTES_INPUT = "textarea[name='notes']"

ACTIVE_CHECKBOX = "input[name='active']"


# =========================
# Buttons / Messages
# =========================

SAVE_BUTTON = "button[type='submit']"

SUCCESS_MESSAGE = ".toast-success"
ERROR_MESSAGE = ".toast-error"