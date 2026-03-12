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

async function searchGoogle(query) {
  const urls = new Set();
  
  const searchEngines = [
    { url: `https://search.brave.com/search?q=${encodeURIComponent(query)}` },
    { url: `https://www.bing.com/search?q=${encodeURIComponent(query)}&count=50` },
  ];

  for (const engine of searchEngines) {
    try {
      const res = await axios.get(engine.url, {
        headers: { 
          'User-Agent': USER_AGENT,
          'Accept': 'text/html,application/xhtml+xml',
          'Accept-Language': 'en-US,en;q=0.9'
        },
        timeout: 15000
      });

      const $ = cheerio.load(res.data);
      
      $('a').each((i, el) => {
        const href = $(el).attr('href');
        if (href && (href.startsWith('http') || href.startsWith('//'))) {
          if (href.startsWith('//')) href = 'https:' + href;
          urls.add(href);
        }
      });
      
      if (urls.size > 0) break;
    } catch (e) {
      console.log('Search error:', e.message);
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
  const { product, country, numLeads } = req.body;

  if (!product || !country) {
    return res.status(400).json({ error: 'Product and country are required' });
  }

  try {
    const queries = SEARCH_SUFFIXES.map(suffix => `${product} ${country} ${suffix}`);

    console.log(`Searching with ${queries.length} queries...`);

    const allUrls = new Set();
    for (const query of queries) {
      const urls = await searchGoogle(query);
      urls.forEach(url => allUrls.add(url));
      await delay(1500);

      if (allUrls.size >= numLeads * 3) break;
    }

    console.log(`Found ${allUrls.size} unique URLs`);

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
    console.log(`Final leads: ${uniqueLeads.length}`);

    const worksheet = XLSX.utils.json_to_sheet(uniqueLeads);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Leads');

    const buffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' });

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename=leads_${product.toLowerCase().replace(/\s+/g, '_')}_${country.toLowerCase()}.xlsx`);
    res.send(buffer);

  } catch (e) {
    console.error('Error:', e);
    res.status(500).json({ error: e.message });
  }
});

const PORT = 3000;
app.listen(PORT, () => console.log(`Server running on http://localhost:${PORT}`));
