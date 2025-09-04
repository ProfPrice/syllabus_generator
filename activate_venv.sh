#!/bin/bash
# Activation script for the syllabus_generator virtual environment

echo "Activating Python virtual environment..."
source venv/bin/activate

echo "Virtual environment activated!"
echo "Python version: $(python --version)"
echo "Pip version: $(pip --version)"
echo ""
echo "To run the syllabus generator:"
echo "  python syllabus_generator.py"
echo ""
echo "To deactivate the virtual environment:"
echo "  deactivate"
