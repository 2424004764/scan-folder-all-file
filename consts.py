# 软件标题
SOFT_NAME = "扫描文件"

# 软件logo地址
LOGO_ICO_PATH = "./logo.ico"

# 敏感词列表
ERROR_WORDS = [
    "木马", "命令", "shell", "测试"
]

# 检查的文件后缀
CHECK_FILE_SUFFIX = ["pdf", "txt", "doc", "docx"]

# 文件扫描列表表头配置
# 切勿轻易修改，除非您知道这么做需要改动什么地方。否则会导致删除异常或者其它问题！
TABLE_HEADER_SUCCESS = "success"
TABLE_HEADER_FAIL = "fail"
TREEVIEW_TABLE_HEADER = {
    TABLE_HEADER_SUCCESS: [
        {
            "key": "index",
            "text": "索引",
            "min_width": 5,
            "width": 5,
            "anchor": "w"
        },
        {
            "key": "file_name",
            "text": "文件名",
            "min_width": 100,
            "width": 200,
            "anchor": "w"
        },
        {
            "key": "opt",
            "text": "忽略文件",
            "min_width": 5,
            "width": 5,
            "anchor": "center"
        }
    ],
    TABLE_HEADER_FAIL: [
        {
            "key": "index",
            "text": "索引",
            "min_width": 5,
            "width": 5,
            "anchor": "w"
        },
        {
            "key": "file_name",
            "text": "文件名",
            "min_width": 100,
            "width": 200,
            "anchor": "w"
        },
        {
            "key": "sensitive_words",
            "text": "命中敏感词",
            "min_width": 100,
            "width": 100,
            "anchor": "center"
        },
        {
            "key": "opt",
            "text": "忽略文件",
            "min_width": 5,
            "width": 5,
            "anchor": "center"
        },
        {
            "key": "delete",
            "text": "从文件内容删除",
            "min_width": 40,
            "width": 40,
            "anchor": "center"
        },
    ]
}
