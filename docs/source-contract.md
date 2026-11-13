# 协议修订资料
`/agreement.docx` 是真实ZIP格式的OOXML包，含正文、编号、脚注和部件关系。`kind=unpaired` 将移动目标改为不匹配的修订ID。正文包含插入、删除、跨段移动以及表格单元格；numId=4从5开始编号。`kind=renumbered` 仅把起始编号改为9，正文不动，用于演示内容未变的重新编号。
当前查看器列出部件并抽取所有w:t节点，不决定接受或拒绝修订，也不计算条款编号。采集必须依据OOXML命名空间及关系解析，不应把查看器的文本数组作为法律正文。

## 条款采集契约
`/collect?view=accept|reject|pending` 输出条款集合，view 必填：`accept` 为拟生效文本（插入与移入生效、删除与移出移除），`reject` 为修订前原文，`pending` 并录双方并标注修订ID、作者与移动配对键。条款身份为 `部件#w14:paraId`，编号依据 `number.basis`（numId、abstractNumId、ilvl、start、startOverride、numFmt、lvlText）可追溯。页眉页脚列入 `excludedParts`，脚注随条款 `footnotes` 对应。`review` 数组列出需人工复核事项（移动配对不完整、关系缺失、编号或脚注无法解析等），采集器不猜测归属。
