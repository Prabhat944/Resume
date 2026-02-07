#!/bin/bash

set -e  # Exit on error

# Colors for output
GREEN='\033[0;32m'
BLUE='\033[0;34m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
NC='\033[0m' # No Color

# Get the directory where the script is located
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
cd "$SCRIPT_DIR"

# Activate virtual environment if it exists
if [ -d ".venv" ]; then
    echo -e "${BLUE}Activating virtual environment...${NC}"
    source .venv/bin/activate
elif [ -d "venv" ]; then
    echo -e "${BLUE}Activating virtual environment...${NC}"
    source venv/bin/activate
else
    echo -e "${YELLOW}Warning: Virtual environment not found. Using system Python.${NC}"
fi

# Check if LibreOffice is installed
LIBREOFFICE_PATH="/Applications/LibreOffice.app/Contents/MacOS/soffice"
if [ ! -f "$LIBREOFFICE_PATH" ]; then
    echo -e "${RED}Error: LibreOffice not found at $LIBREOFFICE_PATH${NC}"
    echo -e "${YELLOW}Please install LibreOffice or update the path in this script.${NC}"
    exit 1
fi

# Function to generate DOCX from Python script
generate_docx() {
    local py_file=$1
    local filename=$(basename "$py_file" .py)
    
    echo -e "${BLUE}Generating DOCX from: $py_file${NC}"
    
    if python3 "$py_file" 2>/dev/null; then
        echo -e "${GREEN}✓ DOCX generated successfully${NC}"
        return 0
    else
        echo -e "${RED}✗ Failed to generate DOCX from $py_file${NC}"
        return 1
    fi
}

# Function to convert DOCX to PDF
convert_to_pdf() {
    local docx_file=$1
    
    if [ ! -f "$docx_file" ]; then
        echo -e "${YELLOW}Warning: $docx_file not found, skipping PDF conversion${NC}"
        return 1
    fi
    
    echo -e "${BLUE}Converting to PDF: $docx_file${NC}"
    
    if "$LIBREOFFICE_PATH" --headless --convert-to pdf --outdir . "$docx_file" > /dev/null 2>&1; then
        echo -e "${GREEN}✓ PDF generated successfully${NC}"
        return 0
    else
        echo -e "${RED}✗ Failed to convert $docx_file to PDF${NC}"
        return 1
    fi
}

# Main execution
echo -e "${BLUE}========================================${NC}"
echo -e "${BLUE}Resume Generation Script${NC}"
echo -e "${BLUE}========================================${NC}"
echo ""

# Find all Python files that generate resumes (excluding venv and common utility files)
PY_FILES=$(find . -maxdepth 1 -name "create_resume*.py" -type f | sort)

if [ -z "$PY_FILES" ]; then
    echo -e "${RED}No resume generation Python files found!${NC}"
    exit 1
fi

echo -e "${BLUE}Found Python files:${NC}"
echo "$PY_FILES" | sed 's|^\./||' | nl
echo ""

# Track success/failure
SUCCESS_COUNT=0
FAIL_COUNT=0
GENERATED_DOCX=()

# Generate DOCX files
echo -e "${BLUE}--- Generating DOCX files ---${NC}"
for py_file in $PY_FILES; do
    if generate_docx "$py_file"; then
        SUCCESS_COUNT=$((SUCCESS_COUNT + 1))
        # Extract the output filename from the Python script
        # This is a simple approach - you may need to adjust based on your script structure
        filename=$(basename "$py_file" .py)
        # Try common output filename patterns
        if [ -f "Prabhat_Kumar_Resume_TwoColumn.docx" ] && [[ "$py_file" == *"twocolumn"* ]] && [[ "$py_file" != *"twocolumn2"* ]] && [[ "$py_file" != *"twocolumn3"* ]]; then
            GENERATED_DOCX+=("Prabhat_Kumar_Resume_TwoColumn.docx")
        elif [ -f "Prabhat_Kumar_Resume_AppDeveloper.docx" ] && [[ "$py_file" == *"app_developer"* ]]; then
            GENERATED_DOCX+=("Prabhat_Kumar_Resume_AppDeveloper.docx")
        elif [ -f "Prabhat_Kumar_Resume_WebDeveloper.docx" ] && [[ "$py_file" == *"twocolumn3"* ]]; then
            GENERATED_DOCX+=("Prabhat_Kumar_Resume_WebDeveloper.docx")
        elif [ -f "Rajnish-Kumar.docx" ] && [[ "$py_file" == *"rajnish"* ]]; then
            GENERATED_DOCX+=("Rajnish-Kumar.docx")
        elif [ -f "Prabhat-Kumar-React.docx" ] && [[ "$py_file" == *"react_template"* ]]; then
            GENERATED_DOCX+=("Prabhat-Kumar-React.docx")
        fi
    else
        FAIL_COUNT=$((FAIL_COUNT + 1))
    fi
    echo ""
done

# Convert DOCX to PDF
echo -e "${BLUE}--- Converting DOCX to PDF ---${NC}"
for docx_file in "${GENERATED_DOCX[@]}"; do
    if [ -f "$docx_file" ]; then
        convert_to_pdf "$docx_file"
        echo ""
    fi
done

# Also try to find and convert any DOCX files that might have been generated
echo -e "${BLUE}--- Checking for additional DOCX files ---${NC}"
ADDITIONAL_DOCX=$(find . -maxdepth 1 \( -name "Prabhat_Kumar_Resume*.docx" -o -name "Rajnish*.docx" -o -name "Prabhat-Kumar-React.docx" \) -type f | sort)
for docx_file in $ADDITIONAL_DOCX; do
    filename=$(basename "$docx_file")
    if [[ ! " ${GENERATED_DOCX[@]} " =~ " ${filename} " ]]; then
        echo -e "${YELLOW}Found additional DOCX: $filename${NC}"
        convert_to_pdf "$docx_file"
        echo ""
    fi
done

# Summary
echo -e "${BLUE}========================================${NC}"
echo -e "${BLUE}Summary${NC}"
echo -e "${BLUE}========================================${NC}"
echo -e "${GREEN}Successfully processed: $SUCCESS_COUNT file(s)${NC}"
if [ $FAIL_COUNT -gt 0 ]; then
    echo -e "${RED}Failed: $FAIL_COUNT file(s)${NC}"
fi
echo ""

# List generated files
echo -e "${BLUE}Generated files:${NC}"
echo -e "${GREEN}DOCX files:${NC}"
(ls -lh Prabhat_Kumar_Resume*.docx 2>/dev/null; ls -lh Rajnish*.docx 2>/dev/null; ls -lh Prabhat-Kumar-React.docx 2>/dev/null) | awk '{print "  " $9 " (" $5 ")"}' || echo "  None"
echo ""
echo -e "${GREEN}PDF files:${NC}"
(ls -lh Prabhat_Kumar_Resume*.pdf 2>/dev/null; ls -lh Rajnish*.pdf 2>/dev/null; ls -lh Prabhat-Kumar-React.pdf 2>/dev/null) | awk '{print "  " $9 " (" $5 ")"}' || echo "  None"
echo ""

echo -e "${GREEN}All done!${NC}"
