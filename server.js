const express = require('express');
const axios = require('axios');
const cheerio = require('cheerio');
const XLSX = require('xlsx');

const app = express();
app.use(express.json());
app.use(express.static('public'));

const SEARCH_SUFFIXES = ['importer', 'distributor', 'wholesale', 'trader', 'supplier', 'buyer'];
const CONTACT_PATHS = ['/contact', '/contact-us', '/about-us', '/about', '/company', '/', '/en/contact', '/contact.html', '/contact-us.html', '/about.html'];

const EMAIL_REGEX = /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g;
const PHONE_REGEX = /(?:\+?\d{1,3}[-.\s]?)?\(?\d{2,4}\)?[-.\s]?\d{2,4}[-.\s]?\d{2,4}/g;

const USER_AGENT = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36';

const DISPOSABLE_DOMAINS = [
  'tempmail.com', '10minutemail.com', 'guerrillamail.com', 'mailinator.com',
  'throwaway.email', 'fakeinbox.com', 'trashmail.com', 'yopmail.com',
  'getnada.com', 'mintemail.com', 'sharklasers.com', 'spam4.me'
];

const BLOCKED_DOMAINS = [
  'google', 'facebook', 'youtube', 'twitter', 'linkedin', 'instagram',
  'amazon', 'ebay', 'bing', 'yahoo', 'duckduckgo', 'yelp', 'yellowpages',
  'pinterest', 'reddit', 'wikipedia', 'wordpress', 'blogspot'
];

function delay(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

function extractDomain(url) {
  try {
    const parsed = new URL(url);
    return parsed.hostname.replace(/^www\./, '').toLowerCase();
  } catch {
    return '';
  }
}

function isValidDomain(domain) {
  if (!domain) return false;
  const lower = domain.toLowerCase();
  return !BLOCKED_DOMAINS.some(b => lower.includes(b));
}

function isValidEmail(email) {
  if (!email) return false;
  const cleaned = email.toLowerCase().trim();
  
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  if (!emailRegex.test(cleaned)) return false;
  
  const parts = cleaned.split('@');
  if (parts.length !== 2) return false;
  
  const domain = parts[1];
  if (DISPOSABLE_DOMAINS.includes(domain)) return false;
  if (domain.includes('example.com') || domain.includes('test.com')) return false;
  
  return true;
}

function isValidPhone(phone) {
  if (!phone) return false;
  
  const digits = phone.replace(/\D/g, '');
  if (digits.length < 6 || digits.length > 15) return false;
  
  const invalidPatterns = ['0000000000', '1111111111', '1234567890', '9876543210'];
  if (invalidPatterns.includes(digits)) return false;
  
  return true;
}

function cleanEmail(email) {
  if (!isValidEmail(email)) return '';
  return email.toLowerCase().trim();
}

function cleanPhone(phone) {
  if (!isValidPhone(phone)) return '';
  return phone.replace(/\s+/g, ' ').trim();
}

function cleanCompany(company) {
  if (!company) return '';
  const cleaned = company
    .replace(/[^\w\s&'.-]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  
  if (cleaned.length < 2 || cleaned.length > 80) return '';
  if (cleaned.toLowerCase().includes('contact') || cleaned.toLowerCase().includes('about')) return '';
  
  return cleaned.slice(0, 80);
}

function cleanUrl(url) {
  if (!url) return '';
  try {
    const parsed = new URL(url);
    if (!isValidDomain(parsed.hostname)) return '';
    return parsed.href.replace(/\/$/, '').split('?')[0];
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

function cleanLeads(leads) {
  return leads.map(lead => ({
    Company: cleanCompany(lead.Company) || extractDomain(lead.Website),
    Website: cleanUrl(lead.Website),
    Email: cleanEmail(lead.Email),
    Phone: cleanPhone(lead.Phone),
    Country: lead.Country,
    Product: lead.Product
  })).filter(lead => 
    lead.Website
  );
}

async function searchBusiness(query) {
  const urls = new Set();
  
  const headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
    'Accept-Language': 'en-US,en;q=0.9,en-GB;q=0.8',
    'Accept-Encoding': 'gzip, deflate, br',
    'Connection': 'keep-alive',
    'Cache-Control': 'max-age=0',
    'Sec-Fetch-Dest': 'document',
    'Sec-Fetch-Mode': 'navigate',
    'Sec-Fetch-Site': 'none',
    'Sec-Fetch-User': '?1',
    'Upgrade-Insecure-Requests': '1'
  };

  const engineConfigs = [
    { name: 'bing', url: `https://www.bing.com/search?q=${encodeURIComponent(query)}&first=0&count=50`, delay: 2000 },
    { name: 'brave', url: `https://search.brave.com/search?q=${encodeURIComponent(query)}&num=50`, delay: 3000 },
    { name: 'yahoo', url: `https://search.yahoo.com/search?p=${encodeURIComponent(query)}&n=50`, delay: 2000 },
    { name: 'duckduckgo', url: `https://html.duckduckgo.com/html/?q=${encodeURIComponent(query)}`, delay: 2000 },
    { name: 'startpage', url: `https://www.startpage.com/do/search?q=${encodeURIComponent(query)}&lui=english`, delay: 2500 },
    { name: 'qwant', url: `https://www.qwant.com/?q=${encodeURIComponent(query)}&num=30`, delay: 2500 }
  ];

  for (const engine of engineConfigs) {
    if (urls.size >= 30) break;
    
    try {
      console.log(`Trying ${engine.name}...`);
      await delay(engine.delay);
      
      const res = await axios.get(engine.url, { 
        headers, 
        timeout: 25000,
        validateStatus: (status) => status < 500
      });
      
      if (res.status === 429) {
        console.log(`Rate limited on ${engine.name}, waiting longer...`);
        await delay(5000);
        continue;
      }
      
      if (res.status !== 200 || res.data.length < 1500) {
        console.log(`${engine.name} returned empty or blocked`);
        continue;
      }
      
      const $ = cheerio.load(res.data);
      
      let found = 0;
      $('a').each((i, el) => {
        if (found >= 50) return;
        
        const href = $(el).attr('href');
        if (!href) return;
        
        let cleanUrl = href;
        
        if (href.includes('/url?') || href.includes('redirect.html')) {
          try {
            const urlParams = new URLSearchParams(href.split('?')[1] || '');
            const direct = urlParams.get('url') || urlParams.get('q') || href.split('=')[1];
            if (direct && direct.startsWith('http')) cleanUrl = decodeURIComponent(direct);
          } catch {}
        }
        
        if (cleanUrl.startsWith('//')) cleanUrl = 'https:' + cleanUrl;
        if (!cleanUrl.startsWith('http') || !cleanUrl.includes('.')) return;
        
        try {
          const domain = extractDomain(cleanUrl);
          if (domain && isValidDomain(domain) && !domain.includes('search') && !domain.includes('bing') && !domain.includes('yahoo') && !domain.includes('brave')) {
            const cleaned = cleanUrl.split('&')[0].split('?')[0].split('#')[0];
            if (cleaned.length > 10) {
              urls.add(cleaned);
              found++;
            }
          }
        } catch {}
      });
      
      if (urls.size > 0) {
        console.log(`Found ${urls.size} URLs from ${engine.name}`);
        if (urls.size >= 20) break;
      }
    } catch (e) {
      const msg = e.message || '';
      if (msg.includes('429') || msg.includes('rate') || msg.includes('Too Many')) {
        console.log(`Rate limited on ${engine.name}, waiting...`);
        await delay(5000);
      } else {
        console.log(`${engine.name} error: ${msg.substring(0, 50)}`);
      }
    }
  }

  console.log(`Total URLs found: ${urls.size}`);
  return Array.from(urls).slice(0, 60);
}

function extractInfo(html) {
  const $ = cheerio.load(html);
  const text = $('body').text();
  
  const emails = text.match(EMAIL_REGEX) || [];
  const validEmail = emails.find(e => isValidEmail(e)) || '';
  
  const phones = text.match(PHONE_REGEX) || [];
  const validPhone = phones.find(p => isValidPhone(p)) || '';
  
  let company = '';
  const title = $('title').text().split('|')[0].split('-')[0].trim();
  const ogSite = $('meta[property="og:site_name"]').attr('content');
  const h1 = $('h1').first().text().split('|')[0].split('-')[0].trim();
  
  for (const text of [ogSite, title, h1]) {
    if (text && text.length > 2 && text.length < 60) {
      company = text;
      break;
    }
  }
  
  return { company, email: validEmail, phone: validPhone };
}

async function visitWebsite(baseUrl) {
  let result = { company: '', website: baseUrl, email: '', phone: '' };
  
  for (const path of CONTACT_PATHS) {
    if (result.email && result.phone && result.company) break;
    
    const url = baseUrl.replace(/\/$/, '') + path;
    try {
      const res = await axios.get(url, {
        headers: { 'User-Agent': USER_AGENT },
        timeout: 8000,
        maxRedirects: 3,
        validateStatus: (status) => status < 500
      });
      
      if (res.status === 200 && res.data.length > 100) {
        const info = extractInfo(res.data);
        
        if (!result.email && info.email) result.email = info.email;
        if (!result.phone && info.phone) result.phone = info.phone;
        if (!result.company && info.company) result.company = info.company;
      }
    } catch {
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
    console.log(`Searching: ${product} ${country}`);

    const allUrls = new Set();
    for (const query of queries) {
      if (allUrls.size >= numLeads * 3) break;
      
      console.log(`\nSearching: ${query}`);
      const urls = await searchBusiness(query);
      urls.forEach(url => allUrls.add(url));
      console.log(`Total so far: ${allUrls.size}`);
      
      if (allUrls.size < numLeads) {
        await delay(5000);
      }
    }

    console.log(`Found ${allUrls.size} URLs`);

    const leads = [];
    const visitedDomains = new Set();

    for (const url of allUrls) {
      if (leads.length >= numLeads) break;

      const domain = extractDomain(url);
      if (visitedDomains.has(domain)) continue;
      if (!isValidDomain(domain)) continue;
      
      visitedDomains.add(domain);
      console.log(`Checking: ${domain}`);

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
          console.log(`Lead: ${domain} - Email: ${!!info.email}, Phone: ${!!info.phone}`);
        }
      } catch (e) {
        console.log(`Error: ${domain} - ${e.message}`);
      }

      await delay(1500);
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
