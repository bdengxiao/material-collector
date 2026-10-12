# 第三方依赖与分发说明

项目源码采用 MIT；依赖许可彼此独立。源码仓库不附带第三方DLL、Windows字体或Microsoft运行库。依赖通过pip获取；完整锁定清单见 requirements-lock.txt。

| 直接依赖/工具 | 版本 | 许可/用途 |
|---|---|---|
| PySide6 Essentials / shiboken6 | 6.8.3 | Qt for Python 的 LGPLv3/GPLv3 或商业许可选项；GUI |
| openpyxl | 3.1.5 | MIT；XLSX读写 |
| et_xmlfile | 2.0.0 | MIT；openpyxl依赖 |
| pytest | 9.1.1 | MIT；开发测试 |
| PyInstaller | 6.16.0 | GPL及bootloader分发例外；构建工具 |
| pyinstaller-hooks-contrib | 2025.9 | 以其发行包许可为准；打包hooks |

源码引用系统字体名称，未分发微软雅黑字体文件。自行构建EXE时，PyInstaller可能带入额外二进制依赖，以上清单不是二进制包的完整许可证明。

公开二进制前，请逐项审计实际文件：保留许可证与版权声明，按所选Qt LGPL条款提供对应库源码/获取方式和适用构建资料，允许用户替换动态库并遵守其他要求。不要假定仅公开应用源码或附上LICENSE即可满足全部义务。

参考：

- https://doc.qt.io/qtforpython-6/licenses.html
- https://www.qt.io/development/open-source-lgpl-obligations
- https://pyinstaller.org/en/stable/license.html

licenses/保留本机依赖发行包中的相关许可证原文；不会将第三方代码重新许可为MIT。
