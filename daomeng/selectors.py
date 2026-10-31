from typing import Dict

DMKJ_PACKAGE_NAME = "com.jingcai.apps.qualitydev"

SELECTORS: Dict[str, Dict[str, str]] = {
    # 登录相关
    "登录失效确定按钮": {
        "resourceId": f"{DMKJ_PACKAGE_NAME}:id/dialog_btn",
        "text": "确定",
    },
    "登录失效提示": {
        "resourceId": f"{DMKJ_PACKAGE_NAME}:id/dialog_content",
        "text": "登录已失效，请重新登录",
    },
    "登录账号": {"resourceId": f"{DMKJ_PACKAGE_NAME}:id/editphonea"},
    "登录密码": {"resourceId": f"{DMKJ_PACKAGE_NAME}:id/editpwda"},
    "同意协议": {"resourceId": f"{DMKJ_PACKAGE_NAME}:id/checkbox"},
    "登录按钮": {"resourceId": f"{DMKJ_PACKAGE_NAME}:id/tv_login", "text": "登 录"},
    "登录页面标题": {
        "resourceId": f"{DMKJ_PACKAGE_NAME}:id/logintitle",
        "text": "欢迎来到，到梦空间",
    },
    # 首页及活动相关
    "首页": {"resourceId": f"{DMKJ_PACKAGE_NAME}:id/homtabs"},
    "我的": {"resourceId": f"{DMKJ_PACKAGE_NAME}:id/tab_view_title", "text": "我的"},
    "我的活动": {"resourceId": f"{DMKJ_PACKAGE_NAME}:id/tvName", "text": "我的活动"},
    "活动管理tab": {
        "resourceId": f"{DMKJ_PACKAGE_NAME}:id/radio_activity",
        "text": "管理",
    },
    "活动名称": {"resourceId": f"{DMKJ_PACKAGE_NAME}:id/tv_activity_name"},
    "活动搜索框": {
        "resourceId": f"{DMKJ_PACKAGE_NAME}:id/text_searchcontent",
    },
    "活动筛选": {
        "resourceId": f"{DMKJ_PACKAGE_NAME}:id/radio_condition",
        "text": "筛选",
    },
    "活动筛选：已结束": {
        "resourceId": f"{DMKJ_PACKAGE_NAME}:id/title",
        "text": "已结束",
    },
    "活动管理": {"resourceId": f"{DMKJ_PACKAGE_NAME}:id/tv_tag_bb", "text": "管理活动"},
    # 学分相关
    "管理列表": {"resourceId": f"{DMKJ_PACKAGE_NAME}:id/menurv"},
    "学分管理": {"resourceId": f"{DMKJ_PACKAGE_NAME}:id/menutitle", "text": "学分管理"},
    "积分": {"resourceId": f"{DMKJ_PACKAGE_NAME}:id/tv_score"},
    "颁发": {"resourceId": f"{DMKJ_PACKAGE_NAME}:id/tv_next", "text": "颁 发"},
    "人员搜索框": {
        "resourceId": f"{DMKJ_PACKAGE_NAME}:id/text_searchcontent",
    },
    "未获得任何学分": {
        "resourceId": f"{DMKJ_PACKAGE_NAME}:id/radio_luqu",
        "text": "未获得任何学分",
    },
    "已获得其他学分": {
        "resourceId": f"{DMKJ_PACKAGE_NAME}:id/radio_sign", "text": "已获得其他学分",
    },
    "人员信息": {"resourceId": f"{DMKJ_PACKAGE_NAME}:id/title", "index": "0"},
    "颁发1人": {
        "resourceId": f"{DMKJ_PACKAGE_NAME}:id/tv_step",
        "text": "颁发1人",
    },
    "确定": {"resourceId": "android:id/button1", "text": "确定"},
    # 其他
    "无结果": {
        "resourceId": f"{DMKJ_PACKAGE_NAME}:id/tv_empty_",
        "text": "这里暂无数据，逛逛其它吧！",
    },
    "加载": {"resourceId": f"{DMKJ_PACKAGE_NAME}:id/progress"},
    "发分加载": {"resourceId": f"{DMKJ_PACKAGE_NAME}:id/tasks_view"},
    "跳过": {"resourceId": f"{DMKJ_PACKAGE_NAME}:id/tv_toapp"},
    "关闭": {"resourceId": f"{DMKJ_PACKAGE_NAME}:id/tv_step", "text": "关闭"},
    "筛选": {"resourceId": f"{DMKJ_PACKAGE_NAME}:id/btn_screen", "text": "筛选"},
}