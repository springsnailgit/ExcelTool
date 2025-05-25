#!/bin/bash
# Excel表格对比工具 - 简化自适应版启动脚本

echo "======================================"
echo "Excel表格对比工具 - 简化自适应版"
echo "======================================"
echo ""
echo "🎯 自适应特性："
echo "✅ 鼠标拖拽窗口边框即可调整大小"
echo "✅ 比对设置中两个功能模块等比例自适应"
echo "✅ 输出目录区域始终保持可操作"
echo "✅ 所有组件跟随外边框自动调整"
echo "✅ 最小尺寸保护：900x700"
echo ""

# 检查Python环境
if command -v python3 &> /dev/null; then
    PYTHON_CMD="python3"
elif command -v python &> /dev/null; then
    PYTHON_CMD="python"
else
    echo "❌ 错误：未找到Python环境"
    exit 1
fi

echo "🐍 Python环境: $($PYTHON_CMD --version)"

# 检查主程序文件
if [ ! -f "Excel_Tool.py" ]; then
    echo "❌ 错误：未找到主程序文件"
    exit 1
fi

# 检查虚拟环境
if [ -d ".venv" ]; then
    echo "🔧 激活虚拟环境..."
    source .venv/bin/activate
elif [ -d "venv" ]; then
    echo "🔧 激活虚拟环境..."
    source venv/bin/activate
fi

echo ""
echo "🚀 启动简化自适应版..."
echo "💡 使用提示：直接拖拽窗口边框测试自适应效果"
echo ""

# 启动主程序
$PYTHON_CMD Excel_Tool.py

echo ""
echo "程序已退出" 