@echo off
REM 一键生成示例素材并自动生成演示文稿
cd /d %~dp0..
python scripts\make_samples.py
python main.py --input examples\sample_data -o output\demo.pptx
echo.
echo 完成！打开 output\demo.pptx 查看结果
pause
