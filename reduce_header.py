with open(r'C:\Users\Stom\Documents\tradeflow\webapp\static\style.css', 'r') as f:
    content = f.read()

old = '''.header-emblem-img {
  height: 52px;
  margin: 0 auto;
  display: block;
}

.header-text {
  flex: 1 1 auto;
  border-left: none;
  border-top: 2px solid var(--red);
  padding: 16px 24px 0;
  display: block;
  text-align: center;
}

.header-tagline {
  margin: 0;
  font-size: 20px;
  font-weight: 700;
  color: var(--black);
  line-height: 1.3;
  letter-spacing: -0.02em;
}

.header-inner {
  max-width: 1200px;
  margin: 0 auto;
  padding: 0;
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  text-align: center;
  gap: 0;
  min-height: 80px;
}'''

new = '''.header-emblem-img {
  height: 44px;
  margin: 0 auto;
  display: block;
}

.header-text {
  flex: 1 1 auto;
  border-left: none;
  border-top: 2px solid var(--red);
  padding: 8px 16px 0;
  display: block;
  text-align: center;
}

.header-tagline {
  margin: 0;
  font-size: 17px;
  font-weight: 700;
  color: var(--black);
  line-height: 1.3;
  letter-spacing: -0.02em;
}

.header-inner {
  max-width: 1200px;
  margin: 0 auto;
  padding: 0;
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  text-align: center;
  gap: 0;
  min-height: 56px;
}'''

if old in content:
    content = content.replace(old, new)
    with open(r'C:\Users\Stom\Documents\tradeflow\webapp\static\style.css', 'w') as f:
        f.write(content)
    print('Replacement successful')
else:
    print('Old string not found')
    # Debug: print first 200 chars
    print(content[:500])