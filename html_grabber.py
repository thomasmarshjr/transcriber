import urllib2, html2text

def remove_non_ascii_1(text):
    return ''.join(i for i in text if ord(i)<128)

response = urllib2.urlopen('https://www.sec.gov/Archives/edgar/data/1326801/000132680116000067/fb-3312016x10q.htm')
html = response.read()
fixed = html2text.html2text(html)
x = open('test.txt', 'w+')
x.write(remove_non_ascii_1(fixed))
x.close()