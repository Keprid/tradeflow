with open(r'C:\Users\Stom\Documents\tradeflow\webapp\static\index.html', 'r') as f:
    content = f.read()

# Add aria-labels to the 5 missing icons
replacements = [
    # Line 38 area - sliders-h in step 1
    ('<i class="fas fa-sliders-h"></i> Step 1', '<i class="fas fa-sliders-h" aria-label="Choose analysis type"></i> Step 1'),
    # Line 45 - chart-line
    ('<i class="fas fa-chart-line"></i> Export Performance', '<i class="fas fa-chart-line" aria-label="View report analytics"></i> Export Performance'),
    # Line 49 - box
    ('<i class="fas fa-box"></i> Trade Brief</label>', '<i class="fas fa-box" aria-label="Product profile mode"></i> Trade Brief</label>'),
    # Line 53 - chess
    ('<i class="fas fa-chess"></i> Competitor Analysis</label>', '<i class="fas fa-chess" aria-label="Competitor analysis mode"></i> Competitor Analysis</label>'),
    # Line 166 - hand-point-down
    ('<i class="fas fa-hand-point-down cta-hand"></i>', '<i class="fas fa-hand-point-down cta-hand" aria-label="Select options"></i>'),
]

for old, new in replacements:
    content = content.replace(old, new)

with open(r'C:\Users\Stom\Documents\tradeflow\webapp\static\index.html', 'w') as f:
    f.write(content)

print('Added 5 missing aria-labels')