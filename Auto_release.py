# 以下为函数main部分-----------------------
import sys
import time

import pandas as pd

from class_release import AutoRelease


def main():

    # 外面需要嵌套一个for循环
    current_release_number = 'RLSE0011779'
    release = AutoRelease(current_release_number)  # 生成与本release有关的类

    query = "release=bcfe0d66dbe33910bfa850d3f3961908^state=13"

    start_time = time.time()
    release.main_procedure(query)
    end_time = time.time()
    print("Run time:", end_time-start_time, "s")


if __name__ == "__main__":
    main()
