with open(r'C:\Users\Stom\Documents\tradeflow\webapp\static\index.html', 'r') as f:
    content = f.read()

icon_labels = {
    'fas fa-sliders-h': 'Choose analysis type',
    'fas fa-chart-line': 'View report analytics',
    'fas fa-box': 'Product profile mode',
    'fas fa-chess': 'Competitor analysis mode',
    'fa-boxes': 'Multiple products mode',
    'fa-concierge-bell': 'Export support',
    'fa-circle-info': 'Information tooltip',
    'fa-calendar-day': 'Review period',
    'fa-flask': 'Chemical analysis',
    'fa-download': 'Download report',
    'fa-hand-point-down': 'Select options',
    'fa-play': 'Generate report',
    'fa-terminal': 'Advanced settings',
    'fa-inbox': 'File upload',
    'fa-spinner fa-spin': 'Processing progress',
    'fa-file-word': 'Word document output',
    'fa-file-archive': 'Excel tables output',
    'fa-file-excel': 'Excel output',
    'fa-xmark': 'Close/remove',
    'fa-check-circle': 'Success confirmation',
    'fa-database': 'Data import',
    'fa-circle-check': 'Validation passed',
    'fa-circle-exclamation': 'Warning/attention',
}

for icon_class, label in icon_labels.items():
    pattern = f'class="fas {icon_class}"'
    replacement = f'class="fas {icon_class}" aria-label="{label}"'
    content = content.replace(pattern, replacement)

with open(r'C:\Users\Stom\Documents\tradeflow\webapp\static\index.html', 'w') as f:
    f.write(content)

print('Done adding aria-labels')