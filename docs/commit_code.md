# 如何向我们提交代码

## 项目使用 poetry 进行管理依赖

https://python-poetry.org/

1, 安装

```shell

# cmd, bash
$ curl -sSL https://install.python-poetry.org | python3 -

# ps
$ (Invoke-WebRequest -Uri https://install.python-poetry.org -UseBasicParsing).Content | py -


```

2, 使用

## 提交代码前使用 pre-commit 格式化代码

https://pre-commit.com/#usage

1, 安装

```shell
pip install pre-commit
```

2, 初始化

```shell
pre-commit install
```


## cz 工具格式化提交内容

https://commitizen-tools.github.io/commitizen/


1, 安装

```shell
poetry add commitizen --dev
```

2, 提交代码

```shell
cz c
```
