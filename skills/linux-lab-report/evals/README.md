# evals

这套评测分两部分：

- **`evals.json`** —— 三个测试用例（prompt + 期望输出 + 19 条期望）。人读的清单。
- **`grade.py`** —— 程序化打分器。读一份生成的 docx，逐条核对那 19 项，输出 `{"expectations": [{"text", "passed", "evidence"}]}`。

`grade.py` 是**权威**：`evals.json` 里 `expectations` 的顺序与 `grade.py` 中 `add()` 的调用顺序一一对应。改了 `grade.py` 就必须同步这份清单——两者曾经漂移成 16 vs 19，而漂移不会让任何工具报错，只有人工逐条比对才发现。

## 跑一次

```bash
pip install python-docx
python evals/grade.py "路径/2512XXXX_某某某_实验报告_6_标题.docx"
python evals/grade.py 报告.docx 结果.json    # 第二参数可选，落盘 JSON
```

输出每条的 PASS/FAIL 与证据，末尾给出 `N/19 passed`。

## 对照实验怎么搭

`grade.py` 只负责打分，不负责跑 agent。要复现「有 skill vs 无 skill」的对照：

1. 每题各跑一次**装了这个 skill**的 agent、一次**不装**的 agent，各自输出到
   `<eval-id>/with_skill/outputs/` 和 `<eval-id>/without_skill/outputs/`。
2. 对每个 docx 跑 `grade.py`，把结果存成 `<config>/outputs/grading.json`。
3. 聚合脚本（skill-creator 的 `scripts/aggregate_benchmark.py`）认的布局是
   `<eval-id>/<config>/run-1/grading.json`，并且要读里面的
   `summary.pass_rate` / `passed` / `failed` / `total`；viewer 认的是顶层
   `expectations` 数组。两个消费者要的形状不一样，所以两份都要写。

### 一个必须避开的坑

跑基线时**不要让 agent 看得到现成的实验报告**。第一次评测就栽在这里：prompt 里带了落盘路径，而那个课件目录下已经有实验报告 1–5，基线 agent 直接读来照抄，结果拿到和装 skill 一样的满分，整个评测失去区分度。跑之前把工作区和课件目录隔离，或者把已有报告移出 agent 可读范围。

## 全局跳过哪些东西

`grade.py` 里有两张豁免名单，都是踩坑后加的，别随手删：

- **`REAL_CMDS`** —— Linux 命令名本身。`xfs_growfs`、`mkfs.xfs`、`ssh-keygen`、`ssh-copy-id` 这些带下划线或点，但它们是工具名，学生改不了，不算违反「文件名短且无下划线」。
- **`SYS_IDENTS`** —— 协议/服务规定的固定名。`sshd_config`、`authorized_keys`、`id_rsa.pub`、`PermitRootLogin`、`ipv4.ignore-auto-dns`、`network-scripts` 等。网络配置和 SSH 那两讲整篇都在跟这类名字打交道，不豁免的话规则会从头红到尾，真正该抓的（学生自造的长文件名）反而被淹掉。

域名同理，用 `\.(com|cn|org|net|edu|gov|io)` 判掉，不用逐个列入白名单。

**教训**：简单正则分不清「学生自造的名字」和「协议规定的名字」，名单式 allowlist 是唯一可靠的做法，别想着用长度或下划线一刀切。
