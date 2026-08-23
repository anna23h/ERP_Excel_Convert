"""Odoo XML-RPC 只读拉数层（2026-08-23 起）。

**只读**。本层不写 ERP —— Safety Stock 回写继续走 `sales_insight` 产的人工导入表，
保留人工过目、出错能中断（沿用 SPEC「第二档 Odoo API 暂不自动化」的判断，只放开读）。
"""
