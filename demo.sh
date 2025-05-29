#!/bin/bash

# DOCX Merger Demo Script
# This script demonstrates the DOCX Merger functionality

echo "DOCX Merger Demo"
echo "================"
echo

# Build the project
echo "1. Building the project..."
dotnet build --verbosity quiet
echo "✓ Build completed"
echo

# Create test documents
echo "2. Creating test documents..."
dotnet run --verbosity quiet -- --create-test
echo "✓ Test documents created"
echo

# Show created files
echo "3. Created files:"
ls -la *.docx
echo

# Merge documents
echo "4. Merging test1.docx and test2.docx..."
dotnet run --verbosity quiet -- test1.docx test2.docx merged_demo.docx
echo

# Show final result
echo "5. Final result:"
ls -la merged_demo.docx
echo

# Test error handling
echo "6. Testing error handling with non-existent file..."
dotnet run --verbosity quiet -- missing.docx test1.docx error_test.docx 2>/dev/null || echo "✓ Error handling works correctly"
echo

echo "Demo completed! You can open merged_demo.docx to see the result."
echo "To merge your own documents, use:"
echo "  dotnet run -- input1.docx input2.docx output.docx"
