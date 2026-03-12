const express = require('express');
const axios = require('axios');
const cheerio = require('cheerio');
const XLSX = require('xlsx');

const app = express();
app.use(express.json());
app.use(express.static('public'));

const SEARCH_SUFFIXES = ['importer', 'distributor', 'wholesale company', 'trader', 'supplier'];
const CONTACT_PATHS = ['/contact', '/contact-us', '/about', '/company', '/'];

const EMAIL_REGEX = /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g;
const PHONE_REGEX = /(?:\+?\d{1,3}[-.\s]?)?\(?\d{2,4}\)?[-.\s]?\d{2,4}[-.\s]?\d{2,4}[-.\s]?\d{2,6}/g;

const USER_AGENT = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36';

function delay(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

function extractDomain(url) {
  try {
    const parsed = new URL(url);
    return parsed.hostname.replace(/^www\./, '');
  } catch {
    return '';
  }
}

function dedupeLeads(leads) {
  const seen = new Set();
  return leads.filter(lead => {
    const key = (lead.Email || '').toLowerCase() || extractDomain(lead.Website);
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function cleanEmail(email) {
  if (!email) return '';
  const cleaned = email.toLowerCase().trim();
  const invalidDomains = ['example.com', 'test.com', 'localhost', 'domain.com', 'sample.com'];
  if (invalidDomains.some(d => cleaned.includes(d))) return '';
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(cleaned)) return '';
  return cleaned;
}

function cleanPhone(phone) {
  if (!phone) return '';
  const cleaned = phone.replace(/[^\d+]/g, '');
  if (cleaned.length < 8 || cleaned.length > 15) return '';
  const invalidPatterns = ['000000', '111111', '123456', '999999'];
  if (invalidPatterns.some(p => cleaned.includes(p))) return '';
  return phone.trim();
}

function cleanCompany(company) {
  if (!company) return '';
  return company
    .replace(/[^\w\s&'-]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 100);
}

function cleanUrl(url) {
  if (!url) return '';
  try {
    const parsed = new URL(url);
    return parsed.href.replace(/\/$/, '');
  } catch {
    return url;
  }
}

function cleanLeads(leads) {
  return leads.map(lead => ({
    Company: cleanCompany(lead.Company) || extractDomain(lead.Website),
    Website: cleanUrl(lead.Website),
    Email: cleanEmail(lead.Email),
    Phone: cleanPhone(lead.Phone),
    Country: lead.Country,
    Product: lead.Product
  })).filter(lead => 
    (lead.Email || lead.Phone) && lead.Website
  );
}

async function searchGoogle(query, searchEngine = 'all') {
  const urls = new Set();
  
  const engineConfigs = {
    brave: [`https://search.brave.com/search?q=${encodeURIComponent(query)}&num=50`],
    bing: [`https://www.bing.com/search?q=${encodeURIComponent(query)}&count=50&setlang=en`],
    yahoo: [`https://search.yahoo.com/search?p=${encodeURIComponent(query)}&n=50`],
    all: [
      `https://search.brave.com/search?q=${encodeURIComponent(query)}&num=50`,
      `https://www.bing.com/search?q=${encodeURIComponent(query)}&count=50&setlang=en`,
      `https://search.yahoo.com/search?p=${encodeURIComponent(query)}&n=50`
    ]
  };

  const engineUrls = engineConfigs[searchEngine] || engineConfigs.all;
  
  const headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8',
    'Accept-Language': 'en-US,en;q=0.9',
    'Accept-Encoding': 'gzip, deflate, br',
    'Connection': 'keep-alive',
    'Upgrade-Insecure-Requests': '1',
    'Sec-Fetch-Dest': 'document',
    'Sec-Fetch-Mode': 'navigate',
    'Sec-Fetch-Site': 'none',
    'Sec-Fetch-User': '?1',
    'Cache-Control': 'max-age=0',
    'sec-ch-ua': '"Not_A Brand";v="8", "Chromium";v="120", "Google Chrome";v="120"',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua-platform': '"Windows"'
  };

  for (const url of engineUrls) {
    try {
      const res = await axios.get(url, {
        headers,
        timeout: 25000
      });

      if (res.data.length < 2000) {
        console.log(`Short response, possible block`);
        continue;
      }

      const $ = cheerio.load(res.data);
      
      $('a').each((i, el) => {
        const href = $(el).attr('href');
        if (href) {
          let cleanUrl = href;
          if (href.startsWith('/url?')) {
            const urlParam = new URLSearchParams(href.split('?')[1]);
            cleanUrl = urlParam.get('url') || href.split('&')[0];
          }
          if (cleanUrl.startsWith('//')) cleanUrl = 'https:' + cleanUrl;
          if (cleanUrl.startsWith('http') && 
              !cleanUrl.includes('search.brave.com') && 
              !cleanUrl.includes('bing.com') &&
              !cleanUrl.includes('yahoo.com') &&
              !cleanUrl.includes('search.yahoo.com') &&
              !cleanUrl.includes('google.com') &&
              !cleanUrl.includes('facebook') &&
              !cleanUrl.includes('youtube') &&
              !cleanUrl.endsWith('.pdf')) {
            try {
              const parsed = new URL(cleanUrl);
              if (parsed.hostname && !parsed.hostname.includes('search')) {
                urls.add(cleanUrl.split('&')[0].split('?')[0]);
              }
            } catch {}
          }
        }
      });
      
      console.log(`Found ${urls.size} URLs from ${url.split('/')[2]}`);
      if (urls.size >= 10) break;
    } catch (e) {
      console.log(`Error: ${e.message}`);
    }
  }
  
  return Array.from(urls).filter(url => 
    url && url.startsWith('http') && 
    !url.includes('duckduckgo') && !url.includes('youtube') &&
    !url.includes('facebook') && !url.includes('linkedin') &&
    !url.includes('twitter') && !url.endsWith('.pdf') &&
    !url.includes('amazon') && !url.includes('ebay')
  );
}

function extractInfo(html, baseUrl) {
  const $ = cheerio.load(html);
  const text = $('body').text();

  const emails = text.match(EMAIL_REGEX) || [];
  const email = emails.find(e => !e.includes('example') && !e.includes('test')) || '';

  const phones = text.match(PHONE_REGEX) || [];
  const phone = phones.find(p => p.replace(/\D/g, '').length >= 8) || '';

  let company = '';
  const ogTitle = $('meta[property="og:site_name"]').attr('content') || '';
  const title = $('title').text() || '';
  const h1 = $('h1').first().text() || '';

  const candidates = [ogTitle, title, h1].filter(Boolean);
  for (const c of candidates) {
    const cleaned = c.split('|')[0].split('-')[0].split('—')[0].trim();
    if (cleaned.length > 2 && cleaned.length < 80 && !cleaned.toLowerCase().includes('contact')) {
      company = cleaned;
      break;
    }
  }

  return { company, email, phone };
}

async function visitWebsite(baseUrl) {
  const result = { company: '', website: baseUrl, email: '', phone: '' };

  for (const path of CONTACT_PATHS) {
    const url = baseUrl.replace(/\/$/, '') + path;

    try {
      const res = await axios.get(url, {
        headers: { 'User-Agent': USER_AGENT },
        timeout: 10000,
        maxRedirects: 3
      });

      const info = extractInfo(res.data, url);

      if (!result.email && info.email) result.email = info.email;
      if (!result.phone && info.phone) result.phone = info.phone;
      if (!result.company && info.company) result.company = info.company;

      if (result.email && result.phone && result.company) break;
    } catch (e) {
      continue;
    }
  }

  return result;
}

app.post('/generate-leads', async (req, res) => {
  const { product, country, companyType, numLeads, searchEngine } = req.body;

  if (!product || !country) {
    return res.status(400).json({ error: 'Product and country are required' });
  }

  const suffixes = companyType && companyType !== 'all' ? [companyType] : SEARCH_SUFFIXES;
  const queries = suffixes.map(suffix => `${product} ${country} ${suffix}`);

  try {
    console.log(`Searching: ${product} ${country} (${searchEngine})...`);

    const allUrls = new Set();
    for (const query of queries) {
      const urls = await searchGoogle(query, searchEngine);
      urls.forEach(url => allUrls.add(url));
      await delay(1500);

      if (allUrls.size >= numLeads * 3) break;
    }

    console.log(`Found ${allUrls.size} URLs`);

    const leads = [];
    const visitedDomains = new Set();

    for (const url of allUrls) {
      if (leads.length >= numLeads) break;

      const domain = extractDomain(url);
      if (visitedDomains.has(domain)) continue;
      visitedDomains.add(domain);

      console.log(`Visiting: ${url}`);

      try {
        const info = await visitWebsite(url);

        if (info.email || info.phone) {
          leads.push({
            Company: info.company || domain,
            Website: info.website,
            Email: info.email,
            Phone: info.phone,
            Country: country,
            Product: product
          });
        }
      } catch (e) {
        console.log(`Error: ${e.message}`);
      }

      await delay(1000);
    }

    const uniqueLeads = dedupeLeads(leads);
    const cleanData = cleanLeads(uniqueLeads);
    console.log(`Final leads: ${cleanData.length}`);

    const worksheet = XLSX.utils.json_to_sheet(cleanData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Leads');

    const buffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' });

    const filename = `leads_${product.toLowerCase().replace(/[^a-z0-9]/g, '_')}_${country.toLowerCase().replace(/[^a-z0-9]/g, '_')}.xlsx`;
    
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    return res.send(buffer);

  } catch (e) {
    console.error('Error:', e);
    res.status(500).json({ error: e.message });
  }
});

const PORT = 3000;
app.listen(PORT, () => console.log(`Server running on http://localhost:${PORT}`));
