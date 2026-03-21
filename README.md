# SC2320 Presentation - Monotone Classification

Lecture notes for group presentation on monotone classification.

## How to Clone

```bash
git clone <repository-url>
cd sc2320-presentation
```

## Setup Instructions

### 1. Set Up Conda Environment

Create a new conda environment for this project:

```bash
conda create -n sc2320 python=3.9
conda activate sc2320
```

### 2. Install MacTeX using Homebrew

Install MacTeX (the distribution of LaTeX for macOS):

```bash
brew install --cask mactex
```

### 3. Install LaTeX Workshop Extension

1. Open VS Code
2. Go to Extensions (Cmd + Shift + X)
3. Search for "LaTeX Workshop"
4. Click Install on the extension by James Yu

## How to Use LaTeX Workshop Extension

### Running the Build

Use the **Build LaTeX project** button in the extension or the green button on the top right hand corner to build the LaTeX document

### Clean Up Files

Use the **Clean up auxilliary files** button in the extension (under Build LaTeX project) to remove auxiliary build files (.aux, .fdb_latexmk, .fls, .log, etc.)

## How to Use the Template

The template files (template.tex and related files) are provided to make document creation easier. They contain pre-built boxes and formatting structures.

**To use the template:**

1. Copy the boxes and styling elements from `template.tex` as needed
2. **Edit your document in `report.tex`** - this is the main file where you'll add your content
3. Customize the copied elements to fit your content

This approach keeps your main work organized while using the template as a reference for consistent formatting.
