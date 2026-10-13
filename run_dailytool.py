"""Điểm vào của DailyTool khi đóng gói thành ứng dụng Windows."""

from __future__ import annotations

import multiprocessing


if __name__ == "__main__":
    multiprocessing.freeze_support()
    from dailytool.desktop import main

    main()
