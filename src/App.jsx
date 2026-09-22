import { useState, useEffect, useMemo } from 'react';
import Papa from 'papaparse';
import * as XLSX from 'xlsx';
import { 
  UploadCloud, 
  ChevronDown, 
  RefreshCw, 
  BarChart2, 
  FileSpreadsheet, 
  LayoutDashboard, 
  AlertCircle, 
  TrendingUp, 
  TrendingDown, 
  AlertTriangle,
  Calendar,
  Clock,
  Sparkles,
  ArrowUpRight,
  ArrowDownRight,
  Info,
  History,
  Save,
  Trash2,
  GitCompare,
  CheckCircle2,
  FolderOpen,
  Plus,
  Bot,
  Cpu,
  Target,
  Rocket,
  ShieldCheck,
  Zap,
  Edit2,
  Copy,
  Check,
  Search,
  ExternalLink,
  Sliders,
  Minus,
  RotateCcw,
  ListChecks,
  X,
  MessageSquare,
  Send
} from 'lucide-react';
import './index.css';

const SNAPSHOTS_STORAGE_KEY = 'cm_analyzer_snapshots_v1';
const CATALOG_STORAGE_KEY = 'cm_analyzer_catalog_v1';
const STRATEGY_STORAGE_KEY = 'cm_analyzer_global_strategy_v1';
const FEEDBACK_STORAGE_KEY = 'cm_feedback_v1';

// Keyword extractor - strips stop words and returns meaningful tokens
const STOP_WORDS = new Set(['the','is','a','an','and','or','but','in','on','at','to','for','of','with','that','this','it','was','are','be','have','has','had','not','as','by','from','we','they','my','our','i','you','can','do','so','if','its']);
const extractKeywords = (text) => {
  return text.toLowerCase()
    .replace(/[^a-z0-9\s]/g, ' ')
    .split(/\s+/)
    .filter(w => w.length > 2 && !STOP_WORDS.has(w));
};


// Default Global Strategy Rules
const DEFAULT_GLOBAL_STRATEGY = {
  defaultObjective: 'growth', // 'launch', 'growth', 'defense', 'profitability'
  targetAcos: 30,
  targetTacos: 12,
  maxBidCeiling: 2.50,
  bleedSpendThreshold: 15.00,
  highAcosThreshold: 40.0,
  lowCtrThreshold: 0.35,
  lowCvrThreshold: 7.0
};

// Strategic Presets for 1-Click Portfolio Optimization
const STRATEGY_PRESETS = [
  {
    id: 'profitability',
    label: 'Strict Margin',
    badge: '🛡️ Strict Margin',
    desc: 'Strict waste cutoff, low bids, 20% Target ACoS',
    config: {
      defaultObjective: 'profitability',
      targetAcos: 20,
      targetTacos: 8,
      bleedSpendThreshold: 10.00,
      maxBidCeiling: 1.50
    }
  },
  {
    id: 'growth',
    label: 'Balanced Scale',
    badge: '⚖️ Balanced Scale',
    desc: 'Moderate expansion & keyword harvesting, 30% Target ACoS',
    config: {
      defaultObjective: 'growth',
      targetAcos: 30,
      targetTacos: 12,
      bleedSpendThreshold: 15.00,
      maxBidCeiling: 2.50
    }
  },
  {
    id: 'aggressive_growth',
    label: 'Max Growth',
    badge: '🚀 Max Growth',
    desc: 'High bids, aggressive rank push, 40% Target ACoS',
    config: {
      defaultObjective: 'growth',
      targetAcos: 40,
      targetTacos: 15,
      bleedSpendThreshold: 25.00,
      maxBidCeiling: 3.50
    }
  },
  {
    id: 'launch',
    label: 'Launch Mode',
    badge: '🎯 Aggressive Launch',
    desc: 'Discovery & Top-of-Search placement, 45% Target ACoS',
    config: {
      defaultObjective: 'launch',
      targetAcos: 45,
      targetTacos: 18,
      bleedSpendThreshold: 30.00,
      maxBidCeiling: 4.00
    }
  },
  {
    id: 'defense',
    label: 'Brand Defense',
    badge: '🏰 Brand Defense',
    desc: 'Protect brand search terms & PDPs, 25% Target ACoS',
    config: {
      defaultObjective: 'defense',
      targetAcos: 25,
      targetTacos: 10,
      bleedSpendThreshold: 15.00,
      maxBidCeiling: 2.00
    }
  }
];

// Utility to extract ASIN from a string
const extractASIN = (campaignName) => {
  const regex = /(B0[A-Z0-9]{8}|[0-9]{9}(X|[0-9]))/i;
  const match = campaignName.match(regex);
  return match ? match[0].toUpperCase() : 'Unknown';
};

// PPC Taxonomy & Jargon Stopwords to filter when extracting product names
const PPC_STOPWORDS = new Set([
  'sp', 'sb', 'sd', 'sponsored', 'products', 'product', 'brands', 'brand', 'display',
  'exact', 'broad', 'phrase', 'auto', 'pat', 'cat', 'category', 'asintargeting', 'targeting',
  'close', 'loose', 'substitutes', 'complements', 'sub', 'comp', 'kw', 'pt', 'skag',
  'tos', 'pp', 'ros', 'topofsearch', 'productpages', 'restofsearch', 'placements',
  'defense', 'offense', 'branded', 'nonbrand', 'generic', 'competitor', 'conquest',
  'launch', 'scale', 'ranking', 'discovery', 'harvest', 'harvesting', 'research',
  'manual', 'automatic', 'bids', 'budget', 'test', 'v1', 'v2', 'v3', 'v4', 'new', 'old',
  'asin', 'campaign', 'adgroup', 'ag', 'camp', 'top', 'search', 'terms', 'st', 'stis',
  'high', 'low', 'mid', 'tier', 'us', 'uk', 'de', 'fr', 'es', 'it', 'ca', 'mx',
  'usd', 'eur', 'gbp', 'aud', 'sponsoredproducts', 'sponsoredbrands', 'sponsoreddisplay'
]);

// Intelligent AI Product Name, Brand, and Category detection from campaign structures & taxonomy
const detectProductContextFromCampaigns = (asin, campaignNames = []) => {
  if (!campaignNames || campaignNames.length === 0) {
    return {
      productName: `Product (${asin})`,
      brand: 'Amazon Catalog',
      category: 'General E-Commerce'
    };
  }

  const candidatePhrases = [];
  const brandCandidates = {};
  const asinRegex = /(B0[A-Z0-9]{8}|[0-9]{9}[0-9X])/i;

  campaignNames.forEach(camp => {
    if (!camp || typeof camp !== 'string') return;
    
    // Split by delimiters
    const segments = camp.split(/[\_\|\-\/\+\:\#\[\]\(\)]+/).map(s => s.trim()).filter(Boolean);

    // Look for brand candidates (often 1st segment if not stopword and not ASIN)
    if (segments.length >= 2) {
      const first = segments[0].replace(/[^a-zA-Z0-9\s]/g, '').trim();
      const firstLower = first.toLowerCase();
      if (
        first.length >= 2 &&
        !PPC_STOPWORDS.has(firstLower) &&
        !first.match(asinRegex) &&
        !first.match(/^[0-9]+$/)
      ) {
        brandCandidates[first] = (brandCandidates[first] || 0) + 1;
      }
    }

    // Filter tokens for descriptive product title
    const tokens = [];
    camp.split(/[\s\_\|\-\/\+\:\#\[\]\(\)\,\.\"]+/).forEach(rawTok => {
      const tok = rawTok.replace(/[^a-zA-Z0-9]/g, '').trim();
      const lower = tok.toLowerCase();
      if (!tok || tok.length < 2) return;

      // Filter out ASIN
      if (tok.match(asinRegex)) return;
      // Filter out PPC stopwords
      if (PPC_STOPWORDS.has(lower)) return;
      // Filter out numbers, percentages, dates, currency
      if (lower.match(/^[0-9]+(\.[0-9]+)?(%|k|x|usd|eur)?$/i)) return;
      if (lower.match(/^(202[0-9]|jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)$/i)) return;

      tokens.push(tok);
    });

    if (tokens.length > 0) {
      candidatePhrases.push(tokens.join(' '));
    }
  });

  // Pick the best / richest descriptive phrase
  let bestTitle = '';
  if (candidatePhrases.length > 0) {
    const sorted = [...candidatePhrases].sort((a, b) => {
      const lenA = a.split(' ').length;
      const lenB = b.split(' ').length;
      if (lenB !== lenA) return lenB - lenA;
      return b.length - a.length;
    });
    bestTitle = sorted[0];
  }

  // Format into Title Case and remove duplicate words
  if (bestTitle && bestTitle.trim().length > 0) {
    bestTitle = bestTitle
      .split(' ')
      .filter((word, idx, arr) => arr.indexOf(word) === idx)
      .map(w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase())
      .join(' ');
  } else {
    bestTitle = `Product (${asin})`;
  }

  // Best brand detection
  let detectedBrand = 'Store Brand';
  const brandEntries = Object.entries(brandCandidates);
  if (brandEntries.length > 0) {
    brandEntries.sort((a, b) => b[1] - a[1]);
    detectedBrand = brandEntries[0][0];
    detectedBrand = detectedBrand.charAt(0).toUpperCase() + detectedBrand.slice(1);
  }

  // Category inference
  let detectedCategory = 'Amazon Marketplace';
  const textToScan = `${bestTitle} ${campaignNames.join(' ')}`.toLowerCase();
  
  if (textToScan.match(/coffee|tea|mug|press|kitchen|pan|pot|knife|blender|spatula|cup|bottle|flask|utensil|cookware|food|baking|air fryer|grill/)) {
    detectedCategory = 'Home & Kitchen';
  } else if (textToScan.match(/wireless|bluetooth|cable|charger|phone|headphone|earbud|speaker|usb|case|audio|adapter|battery|laptop|monitor|gaming/)) {
    detectedCategory = 'Electronics & Accessories';
  } else if (textToScan.match(/shirt|pants|dress|shoes|socks|jacket|hoodie|hat|cotton|apparel|underwear|belt|bag|backpack|leather/)) {
    detectedCategory = 'Clothing, Shoes & Jewelry';
  } else if (textToScan.match(/serum|cream|lotion|shampoo|skincare|soap|brush|oil|beauty|cosmetic|cleanser|hair|makeup|perfume/)) {
    detectedCategory = 'Beauty & Personal Care';
  } else if (textToScan.match(/yoga|fitness|gym|workout|band|mat|dumbbell|sports|exercise|weights|outdoor|camping|hiking|bike/)) {
    detectedCategory = 'Sports & Outdoors';
  } else if (textToScan.match(/pet|dog|cat|leash|collar|bowl|treat|puppy|litter|aquarium|chew/)) {
    detectedCategory = 'Pet Supplies';
  } else if (textToScan.match(/toy|game|kids|baby|puzzle|plush|doll|lego|board game/)) {
    detectedCategory = 'Toys & Games';
  } else if (textToScan.match(/supplement|vitamin|protein|health|cleaner|detergent|wipes|pill|medical|first aid/)) {
    detectedCategory = 'Health & Household';
  } else if (textToScan.match(/desk|pen|notebook|paper|office|chair|planner|marker/)) {
    detectedCategory = 'Office Products';
  } else if (textToScan.match(/tool|drill|saw|wrench|hardware|garden|hose|light|bulb/)) {
    detectedCategory = 'Tools & Home Improvement';
  }

  return {
    productName: bestTitle,
    brand: detectedBrand,
    category: detectedCategory
  };
};

// Utility to parse currency string to float, handling EU/US formats
const parseCurrency = (val) => {
  if (!val) return 0;
  if (typeof val === 'number') return isNaN(val) ? 0 : val;
  let str = val.toString().trim();
  if (!str) return 0;
  
  const lastComma = str.lastIndexOf(',');
  const lastDot = str.lastIndexOf('.');
  if (lastComma > lastDot) {
    str = str.replace(/\./g, '').replace(',', '.');
  } else {
    str = str.replace(/,/g, '');
  }
  
  const parsed = parseFloat(str.replace(/[^0-9.-]+/g, ''));
  return isNaN(parsed) ? 0 : parsed;
};

// Utility to parse varied date formats
const parseRowDate = (val) => {
  if (!val) return null;
  if (val instanceof Date) return isNaN(val.getTime()) ? null : val;
  if (typeof val === 'number') {
    if (val > 30000 && val < 60000) {
      const d = new Date(Math.round((val - 25569) * 86400 * 1000));
      return isNaN(d.getTime()) ? null : d;
    }
  }
  const str = val.toString().trim();
  if (!str) return null;
  
  let parsed = new Date(str);
  if (!isNaN(parsed.getTime())) return parsed;
  
  const parts = str.split(/[\/\-\.]/);
  if (parts.length === 3 && parts[2].length === 4) {
    parsed = new Date(parseInt(parts[2], 10), parseInt(parts[1], 10) - 1, parseInt(parts[0], 10));
    if (!isNaN(parsed.getTime())) return parsed;
  }
  return null;
};

// Generate insights based on campaign metrics and strategy thresholds
const generateSignals = (camp, strategy = {}) => {
  const signals = [];
  const bleedThreshold = strategy.bleedSpendThreshold || 15;
  const targetAcos = strategy.targetAcos || 30;
  const lowCtrThreshold = strategy.lowCtrThreshold || 0.35;
  const ctr = camp.impressions > 0 ? (camp.clicks / camp.impressions) * 100 : 0;
  
  if (camp.spend > bleedThreshold && camp.sales === 0) {
    signals.push({ type: 'danger', icon: <AlertCircle size={14}/>, msg: `Bleeding $${camp.spend.toFixed(2)} with 0 sales. Pause or cut bids.` });
  } else if (camp.acos > targetAcos * 1.3) {
    signals.push({ type: 'warning', icon: <TrendingDown size={14}/>, msg: `ACoS ${camp.acos.toFixed(1)}% exceeds target ${targetAcos}%. Reduce bids.` });
  } else if (camp.acos > 0 && camp.acos <= targetAcos * 0.8) {
    signals.push({ type: 'success', icon: <TrendingUp size={14}/>, msg: `Profitable at ${camp.acos.toFixed(1)}% ACoS. Consider scaling budget.` });
  }

  if (camp.impressions > 2000 && ctr < lowCtrThreshold) {
    signals.push({ type: 'warning', icon: <AlertTriangle size={14}/>, msg: `Low CTR (${ctr.toFixed(2)}%). Improve main image/title.` });
  }

  if (camp.clicks > 15 && camp.orders === 0) {
    signals.push({ type: 'danger', icon: <AlertCircle size={14}/>, msg: `${camp.clicks} clicks, 0 orders. Check listing/price.` });
  }

  return signals;
};

// Autonomous AI Strategy Generator Engine
const generateAIStrategyReport = (asinGroup, catalogItem, globalStrategy, prevData = null) => {
  const asin = asinGroup.asin;
  const productName = catalogItem?.productName || `Product (${asin})`;
  const brand = catalogItem?.brand || 'Brand Catalog';
  const category = catalogItem?.category || 'General E-Commerce';
  const price = catalogItem?.price || 0;
  const audience = catalogItem?.audience || 'Amazon Marketplace Shoppers';
  
  const isCustomOverride = catalogItem?.strategyMode && catalogItem.strategyMode !== 'global';
  const appliedStrategyMode = isCustomOverride ? catalogItem.strategyMode : globalStrategy.defaultObjective;
  const targetAcos = catalogItem?.targetAcos || globalStrategy.targetAcos;
  const targetTacos = catalogItem?.targetTacos || globalStrategy.targetTacos;
  const customDirectives = catalogItem?.customDirectives || '';

  const totalSpend = asinGroup.totalSpend;
  const totalSales = asinGroup.totalSales;
  const acos = asinGroup.acos;
  const orders = asinGroup.totalOrders;
  const clicks = asinGroup.totalClicks;
  const impressions = asinGroup.totalImpressions;
  const ctr = impressions > 0 ? (clicks / impressions) * 100 : 0;
  const cvr = clicks > 0 ? (orders / clicks) * 100 : 0;
  const avgCpc = clicks > 0 ? totalSpend / clicks : 0;

  // Estimated TACoS (assuming organic sales ratio estimate if total sales given)
  const estimatedTacos = totalSales > 0 ? (totalSpend / (totalSales * 1.35)) * 100 : 0;

  // Step-by-step optimization actions
  const actions = [];
  const rationales = [];

  // Strategy Mode Labels
  const modeLabels = {
    launch: 'Launch / Cold Start (Aggressive Discovery & Top-of-Search)',
    growth: 'Growth / Scale (Keyword Harvesting & Targeted Scaling)',
    defense: 'Defense & Brand Protection (Protect PDPs & High-Performing terms)',
    profitability: 'Profitability & Margin Optimization (Strict Waste Negation & Low Bids)',
    custom: 'Custom ASIN Directive Override'
  };

  // 1. Evaluate Bleeders (Spend without sales)
  const bleedingCampaigns = asinGroup.campaigns.filter(c => c.spend >= globalStrategy.bleedSpendThreshold && c.sales === 0);
  if (bleedingCampaigns.length > 0) {
    bleedingCampaigns.forEach(c => {
      const suggestedBid = Math.max(0.20, c.cpc * 0.55).toFixed(2);
      actions.push({
        type: 'negative_or_bid_cut',
        level: 'danger',
        target: c.name,
        action: appliedStrategyMode === 'launch' 
          ? `Lower bid to $${suggestedBid} (-45%) & review search term relevance`
          : `Pause non-converting search terms or slash bid to $${suggestedBid}`,
        impact: `Saves ~$${c.spend.toFixed(2)} wasted spend directly.`
      });
    });
    rationales.push(`Found ${bleedingCampaigns.length} campaign(s) spending over $${globalStrategy.bleedSpendThreshold} with zero purchases.`);
  }

  // 2. Evaluate High ACoS terms vs Strategy Targets
  const highAcosCamps = asinGroup.campaigns.filter(c => c.sales > 0 && c.acos > targetAcos);
  if (highAcosCamps.length > 0) {
    highAcosCamps.forEach(c => {
      const excessRatio = c.acos / targetAcos;
      const reductionPct = Math.min(40, Math.round((1 - (targetAcos / c.acos)) * 100));
      const targetCpc = (c.cpc * (targetAcos / c.acos)).toFixed(2);
      actions.push({
        type: 'bid_decrease',
        level: 'warning',
        target: c.name,
        action: `Reduce bid by ${reductionPct}% (Target CPC: $${targetCpc}) to bring ACoS from ${c.acos.toFixed(1)}% down to ${targetAcos}%`,
        impact: `Preserves conversion volume while realigning ACoS with profitability goals.`
      });
    });
    rationales.push(`Current ACoS (${acos.toFixed(1)}%) exceeds target threshold of ${targetAcos}%. Bid calibration required on ${highAcosCamps.length} campaigns.`);
  }

  // 3. Evaluate Top Performers / Scale Candidates
  const champions = asinGroup.campaigns.filter(c => c.orders >= 2 && c.acos > 0 && c.acos <= targetAcos * 0.85);
  if (champions.length > 0) {
    champions.forEach(c => {
      const suggestedBid = (c.cpc * 1.20).toFixed(2);
      actions.push({
        type: 'scale_winner',
        level: 'success',
        target: c.name,
        action: `Boost Top-of-Search placement (+20%) and increase campaign budget. Target Bid: $${suggestedBid}`,
        impact: `Captures incremental high-converting impression share while operating below target ACoS (${c.acos.toFixed(1)}%).`
      });
    });
    rationales.push(`Identified ${champions.length} highly efficient champion campaigns yielding strong CVR (${cvr.toFixed(1)}%) with low ACoS.`);
  }

  // 4. Keyword Harvesting for Growth & Launch
  if (appliedStrategyMode === 'growth' || appliedStrategyMode === 'launch') {
    actions.push({
      type: 'harvest',
      level: 'info',
      target: `${asin} Auto / Broad Discovery`,
      action: `Harvest search terms generating ≥ 2 orders and promote to Exact Match single-keyword ad groups (SKAGs). Add exact negatives to Auto campaign.`,
      impact: `Prevents bid cannibalization and locks in winning keyword rank.`
    });
    rationales.push(`Under ${modeLabels[appliedStrategyMode]}, automated discovery harvesting maintains aggressive rank growth.`);
  }

  // 5. CTR & Listing Page Diagnostics
  if (impressions > 1500 && ctr < globalStrategy.lowCtrThreshold) {
    actions.push({
      type: 'listing_diagnostic',
      level: 'warning',
      target: `Listing Front-End (${asin})`,
      action: `Optimize Main Image, Badge/Coupon, and Title. Click-through rate (${ctr.toFixed(2)}%) is below standard (${globalStrategy.lowCtrThreshold}%).`,
      impact: `Higher CTR improves Amazon organic rank velocity and reduces effective CPC.`
    });
  }

  // 6. Custom Strategy Directives Override Handling
  if (isCustomOverride && customDirectives) {
    rationales.unshift(`[CUSTOM OVERRIDE ACTIVE]: User instructed: "${customDirectives}". Priority given to custom directives over default portfolio heuristics.`);
    actions.unshift({
      type: 'custom_directive',
      level: 'primary',
      target: `ASIN Custom Instruction Override`,
      action: `Execute user directive: ${customDirectives}`,
      impact: `Direct alignment with seller's unique inventory or ranking target.`
    });
  }

  // Fallback if no specific issues
  if (actions.length === 0) {
    actions.push({
      type: 'maintain',
      level: 'success',
      target: `All Campaigns for ${asin}`,
      action: `Maintain current bids & budgets. Performance is healthy and aligned with target ACoS (${targetAcos}%).`,
      impact: `Stable pacing and steady revenue generation.`
    });
    rationales.push(`Performance is stable within the designated efficiency envelope.`);
  }

  return {
    asin,
    productIdentity: {
      asin,
      productName,
      brand,
      category,
      price,
      audience
    },
    strategyConfig: {
      isCustomOverride,
      appliedStrategyMode,
      modeLabel: modeLabels[appliedStrategyMode] || appliedStrategyMode,
      targetAcos,
      targetTacos,
      customDirectives
    },
    performanceMetrics: {
      spend: totalSpend,
      sales: totalSales,
      acos,
      estimatedTacos,
      orders,
      clicks,
      impressions,
      ctr,
      cvr,
      avgCpc
    },
    actions,
    rationales
  };
};

// Comparative Week-over-Week Strategy Insights Engine
const generateComparativeAIReport = (row, catItem, globalStrategy) => {
  const asin = row.asin;
  const isCustom = catItem?.strategyMode && catItem.strategyMode !== 'global';
  const mode = isCustom ? catItem.strategyMode : globalStrategy.defaultObjective;
  const targetAcos = catItem?.targetAcos || globalStrategy.targetAcos;
  const title = catItem?.productName || (asin === 'Unknown' ? 'Non-ASIN Grouped Campaigns' : `Product (${asin})`);
  const brand = catItem?.brand || 'Store Brand';
  const category = catItem?.category || 'Amazon Catalog';

  const spendA = row.spendA || 0;
  const spendB = row.spendB || 0;
  const salesA = row.salesA || 0;
  const salesB = row.salesB || 0;
  const acosA = row.acosA || 0;
  const acosB = row.acosB || 0;
  const ordersA = row.ordersA || 0;
  const ordersB = row.ordersB || 0;

  const spendDiff = spendB - spendA;
  const spendPct = spendA > 0 ? ((spendDiff) / spendA) * 100 : (spendB > 0 ? 100 : 0);
  const salesDiff = salesB - salesA;
  const salesPct = salesA > 0 ? ((salesDiff) / salesA) * 100 : (salesB > 0 ? 100 : 0);
  const acosDiff = acosB - acosA;
  const ordersDiff = ordersB - ordersA;

  let badge = { type: 'neutral', label: 'Stable Pacing', color: 'var(--accent-color)', icon: 'pacing' };
  let shortSuggestion = '';
  let rationale = '';
  let priority = 3; // 1 = action required (red), 2 = growth/margin opportunity (green), 3 = stable (blue)
  let categoryKey = 'stable'; // 'action_needed', 'winners', 'margin_improved', 'stable'
  const actionDirectives = [];

  if (spendB > 0 && salesB === 0) {
    badge = { type: 'danger', label: 'Bleeding (0 Sales)', color: '#ef4444', icon: 'alert' };
    shortSuggestion = `Zero orders on $${spendB.toFixed(2)} spend. Pause non-converting search terms and cut base bids by 25-35%.`;
    rationale = `In the comparison period, ad spend reached $${spendB.toFixed(2)} with 0 customer purchases. Immediate budget leak.`;
    priority = 1;
    categoryKey = 'action_needed';
    actionDirectives.push(`Audit search term report for ${asin} and add negative exacts for all zero-converting keywords with > 5 clicks.`);
    actionDirectives.push(`Drop default keyword bids to $0.35-$0.50 to arrest bleeding while investigating listing conversion issues.`);
  } else if (acosB > targetAcos && (acosDiff > 5 || (acosA > targetAcos && acosB > targetAcos))) {
    badge = { type: 'danger', label: 'ACoS Spike / High', color: '#ef4444', icon: 'trend_down' };
    shortSuggestion = `ACoS rose to ${acosB.toFixed(1)}% (vs ${acosA.toFixed(1)}% prior, Target ${targetAcos}%). Reduce bids on high-spend keywords.`;
    rationale = `ACoS deteriorated by ${acosDiff > 0 ? '+' : ''}${acosDiff.toFixed(1)}%, now sitting above the target ceiling of ${targetAcos}%.`;
    priority = 1;
    categoryKey = 'action_needed';
    actionDirectives.push(`Lower bids by 15-20% on all keywords operating with ACoS > ${targetAcos}%.`);
    actionDirectives.push(`Adjust top-of-search placement modifier downward if CPC inflated week-over-week.`);
  } else if (salesDiff > 0 && acosB <= targetAcos && acosB > 0) {
    badge = { type: 'success', label: 'Scale Winner', color: '#10b981', icon: 'trend_up' };
    shortSuggestion = `Sales up +${salesPct.toFixed(1)}% at a profitable ${acosB.toFixed(1)}% ACoS. Boost top-of-search bids & campaign budget.`;
    rationale = `Strong conversion acceleration with efficiency comfortably under the ${targetAcos}% target ACoS threshold.`;
    priority = 2;
    categoryKey = 'winners';
    actionDirectives.push(`Increase daily budget by +25-30% on top converting campaigns for ${asin} to avoid losing impression share.`);
    actionDirectives.push(`Boost Top-of-Search placement modifier by +15% on high-converting exact match keywords.`);
  } else if (acosDiff < -2.5 && acosB > 0) {
    badge = { type: 'success', label: 'Margin Improved', color: '#10b981', icon: 'shield' };
    shortSuggestion = `ACoS improved from ${acosA.toFixed(1)}% down to ${acosB.toFixed(1)}%. Maintain optimal bids & harvest search terms.`;
    rationale = `ACoS dropped by ${Math.abs(acosDiff).toFixed(1)}% while generating $${salesB.toFixed(2)} in sales, expanding unit profitability.`;
    priority = 2;
    categoryKey = 'margin_improved';
    actionDirectives.push(`Hold existing bids steady. The current bid pricing aligns with optimal profit margins.`);
    actionDirectives.push(`Harvest profitable search terms into single-keyword ad groups (SKAGs).`);
  } else if (spendDiff < 0 && salesDiff < 0 && Math.abs(salesPct) > 15) {
    badge = { type: 'warning', label: 'Volume Dropping', color: '#f59e0b', icon: 'alert' };
    shortSuggestion = `Spend contraction caused a -${Math.abs(salesPct).toFixed(1)}% sales drop. Check lost impression share due to budget limits.`;
    rationale = `Cutting spend by $${Math.abs(spendDiff).toFixed(2)} contracted total orders from ${ordersA} to ${ordersB}.`;
    priority = 2;
    categoryKey = 'action_needed';
    actionDirectives.push(`Re-evaluate whether budget throttling choked high-performing keyword targets.`);
  } else {
    badge = { type: 'neutral', label: 'Stable Pacing', color: '#3b82f6', icon: 'check' };
    shortSuggestion = `Performance is steady within acceptable efficiency parameters. Continue baseline monitoring.`;
    rationale = `Week-over-week performance delta is within normal standard deviation bounds.`;
    priority = 3;
    categoryKey = 'stable';
    actionDirectives.push(`Keep current campaign settings active.`);
  }

  return {
    asin,
    title,
    brand,
    category,
    mode,
    targetAcos,
    isCustom,
    badge,
    shortSuggestion,
    rationale,
    priority,
    categoryKey,
    actionDirectives,
    spendA, spendB, spendDiff, spendPct,
    salesA, salesB, salesDiff, salesPct,
    acosA, acosB, acosDiff,
    ordersA, ordersB, ordersDiff
  };
};

const generateOverallComparativeSummary = (aggA, aggB, asinReports, globalStrategy) => {
  const spendDiff = aggB.totalSpend - aggA.totalSpend;
  const spendPct = aggA.totalSpend > 0 ? (spendDiff / aggA.totalSpend) * 100 : (aggB.totalSpend > 0 ? 100 : 0);
  const salesDiff = aggB.totalSales - aggA.totalSales;
  const salesPct = aggA.totalSales > 0 ? (salesDiff / aggA.totalSales) * 100 : (aggB.totalSales > 0 ? 100 : 0);
  const acosDiff = aggB.totalAcos - aggA.totalAcos;

  const actionCount = asinReports.filter(r => r.categoryKey === 'action_needed').length;
  const winnerCount = asinReports.filter(r => r.categoryKey === 'winners').length;
  const marginCount = asinReports.filter(r => r.categoryKey === 'margin_improved').length;

  let trendHeadline = '';
  let trendType = 'neutral'; // 'success', 'danger', 'warning', 'neutral'
  let executiveSummary = '';

  if (salesDiff > 0 && acosDiff <= 0.5) {
    trendHeadline = 'Growth & Profitability Expansion: Revenue scaled with strong ACoS efficiency.';
    trendType = 'success';
    executiveSummary = `Total revenue grew by +${(salesDiff).toLocaleString('en-US', {style: 'currency', currency: 'USD'})} (+${salesPct.toFixed(1)}%) while portfolio ACoS maintained efficiency at ${aggB.totalAcos.toFixed(2)}%. Focus this week on scaling top-performing ASINs.`;
  } else if (salesDiff > 0 && acosDiff > 0.5 && aggB.totalAcos <= globalStrategy.targetAcos) {
    trendHeadline = 'Growth Acceleration: Revenue expanded within acceptable portfolio target ACoS.';
    trendType = 'info';
    executiveSummary = `Sales expanded by +${salesPct.toFixed(1)}% with an incremental spend increase. Portfolio ACoS sits safely at ${aggB.totalAcos.toFixed(2)}% (under target ${globalStrategy.targetAcos}%).`;
  } else if (acosDiff > 3) {
    trendHeadline = 'Efficiency Warning: Portfolio ACoS increased week-over-week. Action required.';
    trendType = 'danger';
    executiveSummary = `Overall ACoS increased by +${acosDiff.toFixed(2)}% (from ${aggA.totalAcos.toFixed(2)}% to ${aggB.totalAcos.toFixed(2)}%). Immediate bid cuts recommended on bleeding ASINs.`;
  } else if (spendDiff < 0 && salesDiff < 0) {
    trendHeadline = 'Revenue Contraction: Spend reduction resulted in lower weekly sales volume.';
    trendType = 'warning';
    executiveSummary = `Lowering ad spend by ${(Math.abs(spendDiff)).toLocaleString('en-US', {style: 'currency', currency: 'USD'})} decreased sales by ${(Math.abs(salesDiff)).toLocaleString('en-US', {style: 'currency', currency: 'USD'})}. Verify impression share loss.`;
  } else {
    trendHeadline = 'Stable Pacing: Portfolio metrics remained consistent week-over-week.';
    trendType = 'neutral';
    executiveSummary = `Week-over-week metrics reflect stable pacing across ASINs with standard conversion fluctuations.`;
  }

  return {
    trendHeadline,
    trendType,
    executiveSummary,
    spendDiff, spendPct,
    salesDiff, salesPct,
    acosDiff,
    actionCount,
    winnerCount,
    marginCount,
    totalAsins: asinReports.length
  };
};

function App() {
  const [activeTab, setActiveTab] = useState('upload'); // 'upload', 'dashboard', 'history', 'compare', 'ai_strategy'
  const [rawRecords, setRawRecords] = useState([]);
  const [hasDates, setHasDates] = useState(false);
  const [dateRangeInfo, setDateRangeInfo] = useState(null);
  const [selectedPeriod, setSelectedPeriod] = useState('all');
  const [loading, setLoading] = useState(false);
  const [expandedAsin, setExpandedAsin] = useState(null);
  const [fileName, setFileName] = useState('');
  const [snapshotNameInput, setSnapshotNameInput] = useState('');
  const [savedSuccessMsg, setSavedSuccessMsg] = useState(false);
  const [copiedActionMsg, setCopiedActionMsg] = useState(false);

  // Feedback Modal State
  const [isFeedbackOpen, setIsFeedbackOpen] = useState(false);
  const [feedbackType, setFeedbackType] = useState('ui');
  const [feedbackText, setFeedbackText] = useState('');
  const [feedbackSubmitted, setFeedbackSubmitted] = useState(false);

  // Feedback Storage (persisted to localStorage)
  const [feedbacks, setFeedbacks] = useState(() => {
    try {
      const saved = localStorage.getItem(FEEDBACK_STORAGE_KEY);
      return saved ? JSON.parse(saved) : [];
    } catch { return []; }
  });

  // Persist feedbacks whenever they change
  useEffect(() => {
    try { localStorage.setItem(FEEDBACK_STORAGE_KEY, JSON.stringify(feedbacks)); } catch {}
  }, [feedbacks]);

  // Compute unread feedback count (status === 'open')
  const unreadFeedbackCount = feedbacks.filter(f => f.status === 'open').length;

  // Cluster feedbacks by shared keywords (2+ shared keywords = same cluster)
  const feedbackClusters = useMemo(() => {
    const clusters = [];
    const assigned = new Set();
    feedbacks.forEach((fb, i) => {
      if (assigned.has(i)) return;
      const cluster = [fb];
      assigned.add(i);
      feedbacks.forEach((other, j) => {
        if (i === j || assigned.has(j)) return;
        const shared = fb.keywords.filter(k => other.keywords.includes(k));
        if (shared.length >= 2) {
          cluster.push(other);
          assigned.add(j);
        }
      });
      if (cluster.length >= 2) {
        const allKeywords = cluster.flatMap(f => f.keywords);
        const freq = {};
        allKeywords.forEach(k => { freq[k] = (freq[k] || 0) + 1; });
        const topKeywords = Object.entries(freq).sort((a,b) => b[1]-a[1]).slice(0, 5).map(e => e[0]);
        clusters.push({ items: cluster, topKeywords, count: cluster.length });
      }
    });
    return clusters.sort((a, b) => b.count - a.count);
  }, [feedbacks]);

  // Submit a new feedback entry
  const submitFeedback = () => {
    if (!feedbackText.trim()) return;
    const entry = {
      id: `fb_${Date.now()}`,
      type: feedbackType,
      text: feedbackText.trim(),
      keywords: extractKeywords(feedbackText),
      timestamp: new Date().toISOString(),
      status: 'open'
    };
    setFeedbacks(prev => [entry, ...prev]);
    setFeedbackSubmitted(true);
  };

  // Mark feedback as resolved / reopen
  const toggleFeedbackStatus = (id) => {
    setFeedbacks(prev => prev.map(f => f.id === id ? { ...f, status: f.status === 'open' ? 'resolved' : 'open' } : f));
  };

  // Delete a single feedback
  const deleteFeedback = (id) => {
    setFeedbacks(prev => prev.filter(f => f.id !== id));
  };



  // ASIN Catalog Memory (Product Names, Custom Strategies, etc.)
  const [asinCatalog, setAsinCatalog] = useState(() => {
    try {
      const saved = localStorage.getItem(CATALOG_STORAGE_KEY);
      return saved ? JSON.parse(saved) : {};
    } catch {
      return {};
    }
  });

  // Global Strategy Settings
  const [globalStrategy, setGlobalStrategy] = useState(() => {
    try {
      const saved = localStorage.getItem(STRATEGY_STORAGE_KEY);
      return saved ? JSON.parse(saved) : DEFAULT_GLOBAL_STRATEGY;
    } catch {
      return DEFAULT_GLOBAL_STRATEGY;
    }
  });
  const [strategySavedToast, setStrategySavedToast] = useState(false);

  // Modal / Drawer state for single ASIN AI strategy inspection & editing
  const [selectedAsinForAI, setSelectedAsinForAI] = useState(null);
  const [editingCatalogAsin, setEditingCatalogAsin] = useState(null);
  const [editForm, setEditForm] = useState({
    productName: '',
    brand: '',
    category: '',
    price: '',
    strategyMode: 'global',
    targetAcos: 30,
    customDirectives: ''
  });

  // Persistent History
  const [history, setHistory] = useState(() => {
    try {
      const saved = localStorage.getItem(SNAPSHOTS_STORAGE_KEY);
      return saved ? JSON.parse(saved) : [];
    } catch {
      return [];
    }
  });

  // Comparison State
  const [compareIdA, setCompareIdA] = useState('');
  const [compareIdB, setCompareIdB] = useState('');
  const [compareFilter, setCompareFilter] = useState('all');
  const [selectedCompareAsinForAI, setSelectedCompareAsinForAI] = useState(null);
  const [compareNotes, setCompareNotes] = useState({}); // { [asin]: string }
  const [exportAllRows, setExportAllRows] = useState(true);

  // Bulk Operations State
  const [bulkActions, setBulkActions] = useState({});

  // Persist Catalog
  useEffect(() => {
    try {
      localStorage.setItem(CATALOG_STORAGE_KEY, JSON.stringify(asinCatalog));
    } catch (e) {
      console.error('Failed to save catalog', e);
    }
  }, [asinCatalog]);

  // Persist Global Strategy
  useEffect(() => {
    try {
      localStorage.setItem(STRATEGY_STORAGE_KEY, JSON.stringify(globalStrategy));
    } catch (e) {
      console.error('Failed to save global strategy', e);
    }
  }, [globalStrategy]);

  // Persist Snapshots
  useEffect(() => {
    try {
      localStorage.setItem(SNAPSHOTS_STORAGE_KEY, JSON.stringify(history));
    } catch (e) {
      console.error('Failed to save history to localStorage', e);
    }
  }, [history]);

  const getVal = (row, substrings) => {
    const keys = Object.keys(row);
    // Priority 1: exact match
    for (const sub of substrings) {
      const exact = keys.find(k => k === sub);
      if (exact) return row[exact];
    }
    // Priority 2: key ends with substring (avoids 'costperclick' matching 'cost')
    for (const sub of substrings) {
      const endMatch = keys.find(k => k.endsWith(sub));
      if (endMatch) return row[endMatch];
    }
    // Priority 3: key starts with substring
    for (const sub of substrings) {
      const startMatch = keys.find(k => k.startsWith(sub));
      if (startMatch) return row[startMatch];
    }
    // Fallback: includes
    for (const sub of substrings) {
      const includeMatch = keys.find(k => k.includes(sub));
      if (includeMatch) return row[includeMatch];
    }
    return 0;
  };

  // Extract date range from filename patterns like "Report_2026-08-05_to_2026-08-11" or "Bulk-08.05.2026-08.11.2026"
  const extractDateRangeFromFilename = (name) => {
    if (!name) return null;
    // Pattern: YYYY-MM-DD_to_YYYY-MM-DD or YYYY-MM-DD - YYYY-MM-DD
    const isoPattern = /(\d{4}-\d{2}-\d{2})[_\s\-]+(?:to[_\s\-]+)?(\d{4}-\d{2}-\d{2})/i;
    let m = name.match(isoPattern);
    if (m) return { start: new Date(m[1]), end: new Date(m[2]) };
    // Pattern: MM.DD.YYYY-MM.DD.YYYY or MM/DD/YYYY
    const dotPattern = /(\d{1,2}[\.\/]\d{1,2}[\.\/]\d{4})[_\s\-]+(\d{1,2}[\.\/]\d{1,2}[\.\/]\d{4})/;
    m = name.match(dotPattern);
    if (m) {
      const parse = (s) => { const p = s.split(/[\.\/]/); return new Date(+p[2], +p[0]-1, +p[1]); };
      return { start: parse(m[1]), end: parse(m[2]) };
    }
    return null;
  };

  // Auto-match the closest prior snapshot for WoW comparison
  const autoMatchPreviousSnapshot = (currentFilenameDateRange) => {
    if (!currentFilenameDateRange || history.length === 0) return;
    const curStart = currentFilenameDateRange.start.getTime();
    
    let bestSnap = null;
    let bestDist = Infinity;
    
    history.forEach(snap => {
      if (!snap.dateRangeInfo?.maxDate) return;
      const snapEnd = new Date(snap.dateRangeInfo.maxDate).getTime();
      const dist = curStart - snapEnd;
      if (dist > 0 && dist < bestDist) {
        bestDist = dist;
        bestSnap = snap;
      }
    });
    
    if (bestSnap && bestDist < 14 * 24 * 60 * 60 * 1000) { // Within 14 days
      setCompareIdA(bestSnap.id);
    }
  };

  const processFile = (file) => {
    if (!file) return;
    setLoading(true);
    setFileName(file.name);
    setBulkActions({});
    
    const fileExtension = file.name.split('.').pop().toLowerCase();

    const onDataReady = (data) => {
      try {
        ingestData(data, file.name);
        const fnDateRange = extractDateRangeFromFilename(file.name);
        if (fnDateRange) autoMatchPreviousSnapshot(fnDateRange);
        setActiveTab('dashboard');
      } catch (err) {
        console.error('Data ingestion failed:', err);
        alert('Error processing file data. Please check the file format.');
      } finally {
        setLoading(false);
      }
    };

    if (fileExtension === 'csv') {
      Papa.parse(file, {
        header: true,
        skipEmptyLines: true,
        complete: (results) => onDataReady(results.data),
        error: (err) => {
          console.error('CSV parse error:', err);
          alert('Failed to parse CSV file. Please check the file format.');
          setLoading(false);
        }
      });
    } else if (fileExtension === 'xlsx' || fileExtension === 'xls') {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const workbook = XLSX.read(e.target.result, { type: 'array' });
          const firstSheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[firstSheetName];
          const json = XLSX.utils.sheet_to_json(worksheet);
          onDataReady(json);
        } catch (err) {
          console.error('Excel parse error:', err);
          alert('Failed to parse Excel file. It may be corrupted or password-protected.');
          setLoading(false);
        }
      };
      reader.onerror = () => {
        alert('Failed to read the file.');
        setLoading(false);
      };
      reader.readAsArrayBuffer(file);
    } else {
      alert("Unsupported file type. Please upload CSV or Excel (.xlsx).");
      setLoading(false);
    }
  };

  const handleFileUpload = (event) => {
    processFile(event.target.files[0]);
  };

  // Drag-and-drop handlers
  const [isDragOver, setIsDragOver] = useState(false);
  const handleDragOver = (e) => { e.preventDefault(); e.stopPropagation(); setIsDragOver(true); };
  const handleDragLeave = (e) => { e.preventDefault(); e.stopPropagation(); setIsDragOver(false); };
  const handleDrop = (e) => {
    e.preventDefault(); e.stopPropagation(); setIsDragOver(false);
    const file = e.dataTransfer?.files?.[0];
    if (file) processFile(file);
  };

  const ingestData = (dataRows, currentFileName = '') => {
    let minDate = null;
    let maxDate = null;
    let dateFoundCount = 0;

    const normalized = [];
    const discoveredAsins = new Set();

    dataRows.forEach(rawRow => {
      const row = {};
      Object.keys(rawRow).forEach(key => {
        if (key) {
          row[key.toLowerCase().trim().replace(/\s+/g, '')] = rawRow[key];
        }
      });

      const campaignName = getVal(row, ['campaignname', 'campaign', 'campaña']);
      if (!campaignName || typeof campaignName !== 'string') return;

      const dateRaw = getVal(row, ['date', 'fecha', 'startdate', 'fechainicio', 'day', 'dia']);
      const parsedDate = dateRaw ? parseRowDate(dateRaw) : null;

      if (parsedDate) {
        dateFoundCount++;
        const time = parsedDate.getTime();
        if (!minDate || time < minDate.getTime()) minDate = parsedDate;
        if (!maxDate || time > maxDate.getTime()) maxDate = parsedDate;
      }

      const asin = extractASIN(campaignName);
      if (asin && asin !== 'Unknown') discoveredAsins.add(asin);

      const spend = parseCurrency(getVal(row, ['spend', 'cost', 'totalcost', 'inversion', 'inversión', 'gasto', 'ausgaben', 'dépenses', 'spesa']));
      const sales = parseCurrency(getVal(row, ['sales', 'venta', 'umsatz', 'ventes', 'vendite']));
      const impressions = parseInt(getVal(row, ['impression', 'impresion'])) || 0;
      const clicks = parseInt(getVal(row, ['click', 'clic'])) || 0;
      const orders = parseInt(getVal(row, ['order', 'pedido', 'purchase', 'purchases'])) || 0;
      const budget = parseCurrency(getVal(row, ['budget', 'presupuesto']));

      normalized.push({
        campaignName,
        asin,
        date: parsedDate ? parsedDate.toISOString() : null,
        spend,
        sales,
        impressions,
        clicks,
        orders,
        budget
      });
    });

    // Group campaign names per ASIN for AI semantic title extraction
    const asinToCampaignsMap = {};
    normalized.forEach(row => {
      if (row.asin && row.asin !== 'Unknown') {
        if (!asinToCampaignsMap[row.asin]) asinToCampaignsMap[row.asin] = [];
        if (!asinToCampaignsMap[row.asin].includes(row.campaignName)) {
          asinToCampaignsMap[row.asin].push(row.campaignName);
        }
      }
    });

    // Auto-discover & enrich catalog memory with AI Product Name & Category detection
    setAsinCatalog(prev => {
      const updated = { ...prev };
      discoveredAsins.forEach(asin => {
        const detected = detectProductContextFromCampaigns(asin, asinToCampaignsMap[asin] || []);
        
        if (!updated[asin]) {
          // Initialize intelligent AI-enriched catalog item
          updated[asin] = {
            asin,
            productName: detected.productName,
            brand: detected.brand,
            category: detected.category,
            price: 0,
            strategyMode: 'global',
            targetAcos: globalStrategy.targetAcos,
            customDirectives: '',
            autoDetected: true
          };
        } else if (!updated[asin].productName || updated[asin].productName === `Product ${asin}` || updated[asin].productName === `Product (${asin})`) {
          // Auto-enrich if existing was just default placeholder
          updated[asin].productName = detected.productName;
          if (!updated[asin].brand || updated[asin].brand === 'Store Brand') updated[asin].brand = detected.brand;
          if (!updated[asin].category || updated[asin].category === 'PPC Catalog') updated[asin].category = detected.category;
          updated[asin].autoDetected = true;
        }
      });
      return updated;
    });

    const datesAvailable = dateFoundCount > 0 && maxDate && minDate;
    setHasDates(datesAvailable);

    let defaultName = `Snapshot - ${new Date().toLocaleDateString(undefined, { month: 'short', day: 'numeric', year: 'numeric' })}`;

    if (datesAvailable) {
      const totalDays = Math.ceil((maxDate.getTime() - minDate.getTime()) / (1000 * 60 * 60 * 24)) + 1;
      const minStr = minDate.toLocaleDateString(undefined, { month: 'short', day: 'numeric' });
      const maxStr = maxDate.toLocaleDateString(undefined, { month: 'short', day: 'numeric', year: 'numeric' });
      
      setDateRangeInfo({
        minDate: minDate.toISOString(),
        maxDate: maxDate.toISOString(),
        totalDays,
        minStr,
        maxStr
      });

      defaultName = `Week of ${minStr} – ${maxStr}`;
      setSelectedPeriod(totalDays >= 7 ? 'this_week' : 'all');
    } else {
      setDateRangeInfo(null);
      setSelectedPeriod('all');
    }

    setSnapshotNameInput(defaultName);
    setRawRecords(normalized);
  };

  const filterRecordsByPeriod = (records, period, rangeInfo) => {
    if (!rangeInfo || period === 'all' || records.length === 0) {
      return { current: records, previous: [] };
    }

    const maxMs = new Date(rangeInfo.maxDate).getTime();
    const oneDayMs = 24 * 60 * 60 * 1000;

    let curStartMs = 0;
    let curEndMs = maxMs + oneDayMs;
    let prevStartMs = 0;
    let prevEndMs = 0;

    if (period === 'this_week') {
      curStartMs = maxMs - (7 * oneDayMs);
      prevEndMs = curStartMs;
      prevStartMs = curStartMs - (7 * oneDayMs);
    } else if (period === 'last_week') {
      curEndMs = maxMs - (7 * oneDayMs);
      curStartMs = maxMs - (14 * oneDayMs);
      prevEndMs = curStartMs;
      prevStartMs = curStartMs - (7 * oneDayMs);
    } else if (period === 'last_30_days') {
      curStartMs = maxMs - (30 * oneDayMs);
      prevEndMs = curStartMs;
      prevStartMs = curStartMs - (30 * oneDayMs);
    }

    const current = records.filter(r => {
      if (!r.date) return false;
      const t = new Date(r.date).getTime();
      return t >= curStartMs && t <= curEndMs;
    });

    const previous = records.filter(r => {
      if (!r.date) return false;
      const t = new Date(r.date).getTime();
      return t >= prevStartMs && t < prevEndMs;
    });

    return { current, previous };
  };

  const aggregateRecords = (records) => {
    const grouped = {};
    let totalSpend = 0;
    let totalSales = 0;
    let totalOrders = 0;
    let totalClicks = 0;
    let totalImpressions = 0;

    // Easy Win detection
    const isEasyWin = (name) => {
      const n = (name || '').toUpperCase();
      return /\bEW\b/.test(n) || /EASY\s*WIN/.test(n);
    };

    records.forEach(row => {
      totalSpend += row.spend;
      totalSales += row.sales;
      totalOrders += row.orders;
      totalClicks += row.clicks;
      totalImpressions += row.impressions;

      const asin = row.asin;
      if (!grouped[asin]) {
        grouped[asin] = {
          asin,
          totalSpend: 0,
          totalSales: 0,
          totalOrders: 0,
          totalClicks: 0,
          totalImpressions: 0,
          campaignMap: {}
        };
      }

      grouped[asin].totalSpend += row.spend;
      grouped[asin].totalSales += row.sales;
      grouped[asin].totalOrders += row.orders;
      grouped[asin].totalClicks += row.clicks;
      grouped[asin].totalImpressions += row.impressions;

      if (!grouped[asin].campaignMap[row.campaignName]) {
        grouped[asin].campaignMap[row.campaignName] = {
          name: row.campaignName,
          spend: 0,
          sales: 0,
          impressions: 0,
          clicks: 0,
          orders: 0,
          budget: row.budget,
          isEasyWin: isEasyWin(row.campaignName)
        };
      }

      const camp = grouped[asin].campaignMap[row.campaignName];
      camp.spend += row.spend;
      camp.sales += row.sales;
      camp.impressions += row.impressions;
      camp.clicks += row.clicks;
      camp.orders += row.orders;
      if (row.budget > camp.budget) camp.budget = row.budget;
    });

    const sortedAsins = Object.values(grouped).map(group => {
      const rawAcos = group.totalSales > 0 ? (group.totalSpend / group.totalSales) * 100 : null;
      group.acos = rawAcos ?? 0;
      group.isBleeding = group.totalSales === 0 && group.totalSpend > 0;
      group.roas = group.totalSpend > 0 ? group.totalSales / group.totalSpend : 0;
      group.avgCpc = group.totalClicks > 0 ? group.totalSpend / group.totalClicks : 0;
      group.ctr = group.totalImpressions > 0 ? (group.totalClicks / group.totalImpressions) * 100 : 0;
      group.cvr = group.totalClicks > 0 ? (group.totalOrders / group.totalClicks) * 100 : 0;
      
      // Performance categorization
      const targetAcos = globalStrategy.targetAcos || 30;
      if (group.isBleeding || (rawAcos !== null && rawAcos > targetAcos * 1.2)) {
        group.perfCategory = 'underperformer';
      } else if (rawAcos !== null && rawAcos <= targetAcos * 0.8 && group.totalOrders >= 2) {
        group.perfCategory = 'top_performer';
      } else {
        group.perfCategory = 'stable';
      }
      
      group.campaigns = Object.values(group.campaignMap).map(camp => {
        const acos = camp.sales > 0 ? (camp.spend / camp.sales) * 100 : (camp.spend > 0 ? null : 0);
        const cpc = camp.clicks > 0 ? camp.spend / camp.clicks : 0;
        const campData = { ...camp, acos: acos ?? 0, isBleeding: camp.sales === 0 && camp.spend > 0, cpc };
        campData.signals = generateSignals(campData, globalStrategy);
        return campData;
      }).sort((a, b) => b.spend - a.spend);

      delete group.campaignMap;
      return group;
    }).sort((a, b) => b.totalSpend - a.totalSpend);

    const totalRoas = totalSpend > 0 ? totalSales / totalSpend : 0;
    const totalAvgCpc = totalClicks > 0 ? totalSpend / totalClicks : 0;
    const totalCtr = totalImpressions > 0 ? (totalClicks / totalImpressions) * 100 : 0;
    const totalCvr = totalClicks > 0 ? (totalOrders / totalClicks) * 100 : 0;

    // Build Easy Win summary: group EW campaigns by ASIN
    const ewByAsin = {};
    sortedAsins.forEach(asinGroup => {
      asinGroup.campaigns.forEach(camp => {
        if (!camp.isEasyWin) return;
        if (!ewByAsin[asinGroup.asin]) {
          ewByAsin[asinGroup.asin] = {
            asin: asinGroup.asin,
            spend: 0, sales: 0, orders: 0,
            campaigns: []
          };
        }
        ewByAsin[asinGroup.asin].spend += camp.spend;
        ewByAsin[asinGroup.asin].sales += camp.sales;
        ewByAsin[asinGroup.asin].orders += camp.orders;
        ewByAsin[asinGroup.asin].campaigns.push(camp);
      });
    });

    const easyWinAsins = Object.values(ewByAsin).map(g => ({
      ...g,
      acos: g.sales > 0 ? (g.spend / g.sales) * 100 : (g.spend > 0 ? null : 0),
      isBleeding: g.sales === 0 && g.spend > 0
    })).sort((a, b) => b.spend - a.spend);

    const ewTotalSpend = easyWinAsins.reduce((s, g) => s + g.spend, 0);
    const ewTotalSales = easyWinAsins.reduce((s, g) => s + g.sales, 0);
    const ewTotalAcos = ewTotalSales > 0 ? (ewTotalSpend / ewTotalSales) * 100 : 0;

    return {
      asins: sortedAsins,
      totalSpend,
      totalSales,
      totalOrders,
      totalClicks,
      totalImpressions,
      totalAcos: totalSales > 0 ? (totalSpend / totalSales) * 100 : 0,
      totalRoas,
      totalAvgCpc,
      totalCtr,
      totalCvr,
      easyWin: {
        asins: easyWinAsins,
        totalSpend: ewTotalSpend,
        totalSales: ewTotalSales,
        totalAcos: ewTotalAcos,
        isBleeding: ewTotalSales === 0 && ewTotalSpend > 0
      }
    };
  };



  const dashboardData = useMemo(() => {
    if (rawRecords.length === 0) return null;

    const { current, previous } = filterRecordsByPeriod(rawRecords, selectedPeriod, dateRangeInfo);
    const curAgg = aggregateRecords(current);
    const prevAgg = previous.length > 0 ? aggregateRecords(previous) : null;

    return {
      current: curAgg,
      previous: prevAgg,
      recordCount: current.length
    };
  }, [rawRecords, selectedPeriod, dateRangeInfo, globalStrategy]);

  // PESH Data Optimizer Engine
  const peshData = useMemo(() => {
    if (rawRecords.length === 0) return null;

    let peshRecords = [];
    
    rawRecords.forEach(row => {
      const campName = String(row['Campaign Name'] || row['Campaign'] || row['campaign'] || '').toUpperCase();
      if (!campName.includes('PESH')) return;

      // Extract entity safely
      const entity = String(row['Entity'] || row['Record Type'] || '').trim().toLowerCase();
      // If Bulk File, restrict to Keyword/Product Targeting.
      // If standard Campaign Manager report, we might not have Entity, but we will grab what we can.
      if (entity && !(entity === 'keyword' || entity === 'product targeting')) return;

      const adGroup = row['Ad Group Name'] || row['Ad Group'] || row['ad group'] || '-';
      const keyword = row['Keyword Text'] || row['Keyword'] || row['Product Targeting Expression'] || row['Targeting'] || '-';
      const matchType = row['Match Type'] || row['match type'] || '-';
      
      const bidVal = parseFloat(row['Bid'] || row['Keyword Bid'] || row['Max Bid']) || null;
      
      const clicks = parseFloat(getVal(row, ['Clicks', 'clicks'])) || 0;
      const spend = parseFloat(getVal(row, ['Spend', 'spend'])) || 0;
      const sales = parseFloat(getVal(row, ['Sales', 'sales', '7 Day Total Sales'])) || 0;
      const orders = parseFloat(getVal(row, ['Orders', 'orders', '7 Day Total Orders'])) || 0;
      
      // Skip empty/invalid rows that might slip through bulk files
      if (keyword === '-' && clicks === 0 && spend === 0) return;

      peshRecords.push({
        campaign: campName,
        adGroup,
        keyword,
        matchType,
        bid: bidVal,
        clicks,
        spend,
        sales,
        orders
      });
    });

    if (peshRecords.length === 0) return null;

    const targetAcos = globalStrategy.targetAcos / 100;
    
    let totalSpend = 0;
    let totalSales = 0;
    
    const analyzed = peshRecords.map(row => {
      totalSpend += row.spend;
      totalSales += row.sales;
      
      const cpc = row.clicks > 0 ? row.spend / row.clicks : 0;
      const cvr = row.clicks > 0 ? row.orders / row.clicks : 0;
      const aov = row.orders > 0 ? row.sales / row.orders : 0;
      const acos = row.sales > 0 ? row.spend / row.sales : (row.spend > 0 ? Infinity : 0);
      const suggestedBid = targetAcos * cvr * aov;
      
      let quadrant = '';
      let action = '';
      let rationale = '';

      if (row.sales === 0 && row.spend >= globalStrategy.bleedThreshold) {
        quadrant = 'Bleeder';
        action = 'Pause / Add Negative';
        rationale = `0 sales, spend > $${globalStrategy.bleedThreshold}. Cut waste.`;
      } else if (row.sales > 0 && acos > targetAcos * 1.5) {
        quadrant = 'Bleeder';
        action = 'Lower Bid 30%';
        rationale = `ACOS ${(acos*100).toFixed(1)}% is > 1.5x target.`;
      } else if (row.sales > 0 && acos <= targetAcos && row.orders >= 2) {
        quadrant = 'Top Performer';
        action = 'Increase Bid 15%';
        rationale = 'Strong efficiency. Scale up.';
      } else if ((row.sales > 0 && acos < targetAcos && row.orders < 2) || (row.clicks > 0 && cvr > 0.15 && row.clicks < 10)) {
        quadrant = 'Under-delivering';
        action = 'Increase Bid 25%';
        rationale = 'High potential CVR/ACOS, low volume.';
      } else if (row.sales > 0 && acos > targetAcos) {
        quadrant = 'Mediocre';
        action = 'Target CPC';
        rationale = `ACOS slightly high. Optimize bid.`;
      } else {
        quadrant = 'Observing';
        action = 'Monitor';
        rationale = 'Not enough data.';
      }

      return {
        ...row,
        cpc,
        cvr,
        aov,
        acos,
        suggestedBid,
        quadrant,
        action,
        rationale
      };
    });

    const bleeders = analyzed.filter(r => r.quadrant === 'Bleeder');
    const topPerformers = analyzed.filter(r => r.quadrant === 'Top Performer');

    return {
      records: analyzed,
      totalSpend,
      totalSales,
      totalAcos: totalSales > 0 ? totalSpend / totalSales : (totalSpend > 0 ? Infinity : 0),
      bleeders,
      counts: {
        bleeders: bleeders.length,
        topPerformers: topPerformers.length,
        underDelivering: analyzed.filter(r => r.quadrant === 'Under-delivering').length,
        mediocre: analyzed.filter(r => r.quadrant === 'Mediocre').length,
        total: analyzed.length
      }
    };
  }, [rawRecords, globalStrategy]);

  // Snapshot actions
  const saveCurrentSnapshot = () => {
    if (rawRecords.length === 0) return;
    const name = snapshotNameInput.trim() || `Snapshot ${history.length + 1}`;
    const agg = aggregateRecords(rawRecords);
    
    const newSnapshot = {
      id: 'snap_' + Date.now(),
      name,
      savedAt: new Date().toISOString(),
      fileName,
      hasDates,
      dateRangeInfo,
      summary: {
        totalSpend: agg.totalSpend,
        totalSales: agg.totalSales,
        totalAcos: agg.totalAcos,
        asinCount: agg.asins.length,
        totalOrders: agg.totalOrders
      },
      rawRecords
    };

    setHistory([newSnapshot, ...history]);
    setSavedSuccessMsg(true);
    setTimeout(() => setSavedSuccessMsg(false), 3000);
  };

  const loadSnapshot = (snap) => {
    setRawRecords(snap.rawRecords);
    setHasDates(snap.hasDates);
    setDateRangeInfo(snap.dateRangeInfo);
    setFileName(snap.fileName || snap.name);
    setSnapshotNameInput(snap.name);
    setSelectedPeriod('all');
    setActiveTab('dashboard');
  };

  const deleteSnapshot = (id, e) => {
    e.stopPropagation();
    if (confirm("Are you sure you want to remove this saved snapshot from history?")) {
      setHistory(history.filter(h => h.id !== id));
      if (compareIdA === id) setCompareIdA('');
      if (compareIdB === id) setCompareIdB('');
    }
  };

  const resetData = () => {
    setRawRecords([]);
    setHasDates(false);
    setDateRangeInfo(null);
    setSelectedPeriod('all');
    setExpandedAsin(null);
    setFileName('');
    setActiveTab('upload');
  };

  const toggleAsin = (asin) => {
    setExpandedAsin(expandedAsin === asin ? null : asin);
  };

  const formatCurrency = (val) => new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD' }).format(val);
  const formatPercent = (val) => `${val.toFixed(2)}%`;
  const formatNumber = (val) => new Intl.NumberFormat('en-US').format(val);

  const renderDelta = (curVal, prevVal, isAcos = false) => {
    if (!prevVal || prevVal === 0) return null;
    const diff = curVal - prevVal;
    const pct = (diff / prevVal) * 100;
    if (Math.abs(pct) < 0.01) return null;

    const isPositiveChange = diff > 0;
    const isGood = isAcos ? !isPositiveChange : isPositiveChange;

    return (
      <span className={`delta-badge ${isGood ? 'delta-good' : 'delta-bad'}`} title={`${isPositiveChange ? '+' : ''}${pct.toFixed(1)}% vs previous`}>
        {isPositiveChange ? <ArrowUpRight size={12} /> : <ArrowDownRight size={12} />}
        {Math.abs(pct).toFixed(1)}%
      </span>
    );
  };

  // Catalog Edit modal handlers
  const openCatalogEditor = (asin, e) => {
    if (e) e.stopPropagation();
    const item = asinCatalog[asin] || {
      asin,
      productName: `Product ${asin}`,
      brand: '',
      category: '',
      price: '',
      strategyMode: 'global',
      targetAcos: globalStrategy.targetAcos,
      customDirectives: ''
    };
    setEditForm({
      productName: item.productName || '',
      brand: item.brand || '',
      category: item.category || '',
      price: item.price || '',
      strategyMode: item.strategyMode || 'global',
      targetAcos: item.targetAcos || globalStrategy.targetAcos,
      customDirectives: item.customDirectives || ''
    });
    setEditingCatalogAsin(asin);
  };

  const saveCatalogEntry = () => {
    if (!editingCatalogAsin) return;
    setAsinCatalog(prev => ({
      ...prev,
      [editingCatalogAsin]: {
        ...prev[editingCatalogAsin],
        asin: editingCatalogAsin,
        productName: editForm.productName.trim() || `Product ${editingCatalogAsin}`,
        brand: editForm.brand.trim(),
        category: editForm.category.trim(),
        price: parseFloat(editForm.price) || 0,
        strategyMode: editForm.strategyMode,
        targetAcos: parseFloat(editForm.targetAcos) || globalStrategy.targetAcos,
        customDirectives: editForm.customDirectives.trim(),
        autoDetected: false
      }
    }));
    setEditingCatalogAsin(null);
  };

  // Auto-detect title for currently edited ASIN in modal
  const autoDetectCurrentAsinTitle = () => {
    if (!editingCatalogAsin) return;
    const camps = rawRecords
      .filter(r => r.asin === editingCatalogAsin)
      .map(r => r.campaignName);

    const detected = detectProductContextFromCampaigns(editingCatalogAsin, camps);
    setEditForm(prev => ({
      ...prev,
      productName: detected.productName,
      brand: detected.brand !== 'Store Brand' ? detected.brand : prev.brand || detected.brand,
      category: detected.category !== 'Amazon Marketplace' ? detected.category : prev.category || detected.category
    }));
  };

  // Re-detect product names across entire catalog from rawRecords
  const redetectAllAsinTitles = () => {
    if (rawRecords.length === 0) {
      alert("Please upload campaign data first to extract product names from campaign structures.");
      return;
    }
    const asinToCampaigns = {};
    rawRecords.forEach(r => {
      if (r.asin && r.asin !== 'Unknown') {
        if (!asinToCampaigns[r.asin]) asinToCampaigns[r.asin] = [];
        if (!asinToCampaigns[r.asin].includes(r.campaignName)) {
          asinToCampaigns[r.asin].push(r.campaignName);
        }
      }
    });

    const count = Object.keys(asinToCampaigns).length;
    if (count === 0) {
      alert("No ASINs found in campaign names to extract.");
      return;
    }

    setAsinCatalog(prev => {
      const updated = { ...prev };
      Object.keys(asinToCampaigns).forEach(asin => {
        const detected = detectProductContextFromCampaigns(asin, asinToCampaigns[asin]);
        updated[asin] = {
          ...(updated[asin] || {
            asin,
            price: 0,
            strategyMode: 'global',
            targetAcos: globalStrategy.targetAcos,
            customDirectives: ''
          }),
          productName: detected.productName,
          brand: detected.brand,
          category: detected.category,
          autoDetected: true
        };
      });
      return updated;
    });

    alert(`✨ AI successfully extracted and enriched product names for ${count} ASINs!`);
  };

  // Strategy Helper Handlers
  const applyStrategyPreset = (preset) => {
    setGlobalStrategy(prev => ({
      ...prev,
      ...preset.config
    }));
    setStrategySavedToast(true);
    setTimeout(() => setStrategySavedToast(false), 2500);
  };

  const handleGlobalStrategyChange = (field, value) => {
    setGlobalStrategy(prev => ({
      ...prev,
      [field]: value
    }));
    setStrategySavedToast(true);
    setTimeout(() => setStrategySavedToast(false), 2000);
  };

  const handleGlobalStrategyBlur = (field, fallback) => {
    setGlobalStrategy(prev => {
      const val = prev[field];
      const num = parseFloat(val);
      return {
        ...prev,
        [field]: isNaN(num) || val === '' ? fallback : num
      };
    });
  };

  const adjustNumericStrategy = (field, step, min = 0, max = 100) => {
    setGlobalStrategy(prev => {
      const current = parseFloat(prev[field]) || 0;
      const updated = Math.min(max, Math.max(min, +(current + step).toFixed(2)));
      return {
        ...prev,
        [field]: updated
      };
    });
    setStrategySavedToast(true);
    setTimeout(() => setStrategySavedToast(false), 2000);
  };

  const applyTargetAcosToAllCatalog = () => {
    const targetAcos = parseFloat(globalStrategy.targetAcos) || 30;
    const asins = Object.keys(asinCatalog);
    if (asins.length === 0) {
      alert("No ASINs currently detected in catalog. Upload campaign data first.");
      return;
    }
    setAsinCatalog(prev => {
      const updated = { ...prev };
      asins.forEach(asin => {
        updated[asin] = {
          ...(updated[asin] || {}),
          targetAcos: targetAcos
        };
      });
      return updated;
    });
    setStrategySavedToast(true);
    alert(`✓ Successfully updated Target ACoS to ${targetAcos}% across all ${asins.length} catalog ASINs!`);
    setTimeout(() => setStrategySavedToast(false), 2500);
  };

  const resetGlobalStrategyToDefaults = () => {
    setGlobalStrategy(DEFAULT_GLOBAL_STRATEGY);
    setStrategySavedToast(true);
    setTimeout(() => setStrategySavedToast(false), 2500);
  };

  // Compare Notes Export Handler
  const exportCompareNotes = (rows, includeAll) => {
    if (!comparisonViewData) return;
    const snapAName = comparisonViewData.snapA?.name || 'Week A';
    const snapBName = comparisonViewData.snapB?.name || 'Week B';

    const exportRows = rows.filter(row => includeAll || (compareNotes[row.asin] || '').trim());

    const wsData = [
      ['ASIN', 'Product Name', 'Campaign Names (A)', 'Campaign Names (B)',
       `Spend (${snapAName})`, `Spend (${snapBName})`,
       `Sales (${snapAName})`, `Sales (${snapBName})`,
       `ACoS (${snapAName})`, `ACoS (${snapBName})`,
       'AI Suggestion', 'Your Notes']
    ];

    exportRows.forEach(row => {
      const catItem = asinCatalog[row.asin];
      const productName = catItem?.productName || row.asin;
      const campsA = (row.campaignsA || []).map(c => c.name).join(', ');
      const campsB = (row.campaignsB || []).map(c => c.name).join(', ');
      wsData.push([
        row.asin,
        productName,
        campsA,
        campsB,
        row.spendA,
        row.spendB,
        row.salesA,
        row.salesB,
        row.acosA ? (row.acosA / 100) : 0,
        row.acosB ? (row.acosB / 100) : 0,
        row.aiReport?.shortSuggestion || '',
        compareNotes[row.asin] || ''
      ]);
    });

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(wsData);

    // Format currency and percent columns
    const currencyCols = [4, 5, 6, 7];
    const percentCols = [8, 9];
    wsData.slice(1).forEach((_, ri) => {
      currencyCols.forEach(ci => {
        const cellRef = XLSX.utils.encode_cell({ r: ri + 1, c: ci });
        if (ws[cellRef]) ws[cellRef].z = '$#,##0.00';
      });
      percentCols.forEach(ci => {
        const cellRef = XLSX.utils.encode_cell({ r: ri + 1, c: ci });
        if (ws[cellRef]) ws[cellRef].z = '0.00%';
      });
    });

    ws['!cols'] = [
      { wch: 14 }, { wch: 36 }, { wch: 40 }, { wch: 40 },
      { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 14 },
      { wch: 12 }, { wch: 12 }, { wch: 42 }, { wch: 40 }
    ];

    XLSX.utils.book_append_sheet(wb, ws, 'WoW Comparison Notes');
    XLSX.writeFile(wb, `WoW_Comparison_Notes_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  // Bulk Operations Handlers
  const [bulkNotes, setBulkNotes] = useState({});

  const handleBulkActionChange = (campaignName, action, currentBudget) => {
    setBulkActions(prev => ({
      ...prev,
      [campaignName]: { action, currentBudget }
    }));
  };

  const handleBulkNoteChange = (campaignName, note) => {
    setBulkNotes(prev => ({ ...prev, [campaignName]: note }));
  };

  // AI Auto-Assign: pre-populate bulk actions based on AI signals
  const aiAutoAssignBulkActions = () => {
    if (rawRecords.length === 0) return;
    const campaignsMap = {};
    rawRecords.forEach(r => {
      if (!campaignsMap[r.campaignName]) {
        campaignsMap[r.campaignName] = { name: r.campaignName, spend: 0, sales: 0, budget: r.budget };
      }
      campaignsMap[r.campaignName].spend += r.spend;
      campaignsMap[r.campaignName].sales += r.sales;
      if (r.budget) campaignsMap[r.campaignName].budget = r.budget;
    });

    const newActions = {};
    Object.values(campaignsMap).forEach(camp => {
      const acos = camp.sales > 0 ? (camp.spend / camp.sales) * 100 : (camp.spend > 0 ? 999 : 0);
      if (camp.spend > globalStrategy.bleedSpendThreshold && camp.sales === 0) {
        newActions[camp.name] = { action: 'pause', currentBudget: camp.budget };
      } else if (acos > globalStrategy.targetAcos * 1.2) {
        newActions[camp.name] = { action: 'reduce_20', currentBudget: camp.budget };
      } else if (acos > 0 && acos < globalStrategy.targetAcos * 0.8 && camp.sales > 0) {
        newActions[camp.name] = { action: 'scale_20', currentBudget: camp.budget };
      }
    });
    setBulkActions(prev => ({ ...prev, ...newActions }));
  };

  const getBulkActionSummary = () => {
    let paused = 0, enabled = 0, scaled = 0, reduced = 0;
    Object.values(bulkActions).forEach(({ action }) => {
      if (action === 'pause') paused++;
      else if (action === 'enable') enabled++;
      else if (action === 'scale_20' || action === 'scale_30' || action === 'scale_50') scaled++;
      else if (action === 'reduce_20') reduced++;
    });
    const total = paused + enabled + scaled + reduced;
    return { paused, enabled, scaled, reduced, total };
  };

  const generateBulkFile = () => {
    const operations = [];
    
    Object.entries(bulkActions).forEach(([campaignName, config]) => {
      const { action, currentBudget } = config;
      if (!action || action === 'none') return;
      
      let state = 'Enabled';
      let newBudget = currentBudget || 10;
      
      if (action === 'pause') {
        state = 'Paused';
        newBudget = currentBudget || '';
      } else if (action === 'enable') {
        state = 'Enabled';
      } else if (action === 'scale_20') {
        newBudget = parseFloat((currentBudget * 1.2).toFixed(2));
      } else if (action === 'scale_30') {
        newBudget = parseFloat((currentBudget * 1.3).toFixed(2));
      } else if (action === 'scale_50') {
        newBudget = parseFloat((currentBudget * 1.5).toFixed(2));
      } else if (action === 'reduce_20') {
        newBudget = parseFloat((currentBudget * 0.8).toFixed(2));
      }

      const row = {
        'Campaign Name': campaignName,
        'Entity': 'Campaign',
        'Operation': 'Update',
        'State': state,
        'Daily Budget': newBudget
      };
      
      const note = bulkNotes[campaignName];
      if (note) row['Notes'] = note;
      
      operations.push(row);
    });

    if (operations.length === 0) {
      alert('No bulk actions selected. Please select at least one action (e.g. Pause, Scale) before exporting.');
      return;
    }

    const worksheet = XLSX.utils.json_to_sheet(operations);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Bulk Operations");
    
    const exportFileName = `Amazon_Bulk_Operations_${new Date().toISOString().split('T')[0]}.xlsx`;
    XLSX.writeFile(workbook, exportFileName);
  };

  // Open AI Strategy Audit Drawer
  const openAIModalForAsin = (asinGroup, e) => {
    if (e) e.stopPropagation();
    const catItem = asinCatalog[asinGroup.asin];
    const plan = generateAIStrategyReport(asinGroup, catItem, globalStrategy, dashboardData?.previous);
    setSelectedAsinForAI({ asinGroup, plan });
  };

  const copyAIPlanText = (plan) => {
    const text = `=== AMAZON PPC AI OPTIMIZATION PLAN ===
1. PRODUCT IDENTITY:
   - Product: ${plan.productIdentity.productName}
   - ASIN: ${plan.productIdentity.asin} | Brand: ${plan.productIdentity.brand}
   - Category: ${plan.productIdentity.category}

2. STRATEGY APPLIED:
   - Mode: ${plan.strategyConfig.modeLabel}
   - Target ACoS: ${plan.strategyConfig.targetAcos}% | Current ACoS: ${plan.performanceMetrics.acos.toFixed(2)}%
   - Spend: $${plan.performanceMetrics.spend.toFixed(2)} | Sales: $${plan.performanceMetrics.sales.toFixed(2)} | Orders: ${plan.performanceMetrics.orders}

3. ACTIONABLE OPTIMIZATIONS:
${plan.actions.map((act, i) => `   [Action ${i+1}] ${act.target}: ${act.action}\n      Impact: ${act.impact}`).join('\n\n')}

4. STRATEGIC RATIONALE:
${plan.rationales.map(r => `   • ${r}`).join('\n')}
=======================================`;

    navigator.clipboard.writeText(text);
    setCopiedActionMsg(true);
    setTimeout(() => setCopiedActionMsg(false), 2500);
  };

  const copyCompareAIPlanText = (item) => {
    if (!item) return;
    const { report, snapA, snapB } = item;
    const text = `=======================================
🤖 AI COMPARATIVE STRATEGY DIRECTIVE (WoW)
=======================================
PRODUCT: ${report.title}
ASIN: ${report.asin}
BRAND: ${report.brand} | CATEGORY: ${report.category}
STRATEGY MODE: ${report.mode.toUpperCase()} (Target ACoS: ${report.targetAcos}%)

1. COMPARISON DELTA (${snapA?.name || 'Period A'} ➔ ${snapB?.name || 'Period B'}):
   • Spend: $${report.spendA.toFixed(2)} ➔ $${report.spendB.toFixed(2)} (${report.spendDiff >= 0 ? '+' : ''}${report.spendPct.toFixed(1)}%)
   • Sales: $${report.salesA.toFixed(2)} ➔ $${report.salesB.toFixed(2)} (${report.salesDiff >= 0 ? '+' : ''}${report.salesPct.toFixed(1)}%)
   • ACoS: ${report.acosA.toFixed(2)}% ➔ ${report.acosB.toFixed(2)}% (${report.acosDiff >= 0 ? '+' : ''}${report.acosDiff.toFixed(2)}% shift)
   • Orders: ${report.ordersA} ➔ ${report.ordersB} (${report.ordersDiff >= 0 ? '+' : ''}${report.ordersDiff})

2. AI STRATEGIC DIAGNOSIS:
   • Status: [${report.badge.label.toUpperCase()}]
   • Rationale: ${report.rationale}
   • Suggestion: ${report.shortSuggestion}

3. RECOMMENDED ACTIONS FOR THIS WEEK:
${report.actionDirectives.map((act, i) => `   ${i + 1}. ${act}`).join('\n')}
=======================================`;

    navigator.clipboard.writeText(text);
    setCopiedActionMsg(true);
    setTimeout(() => setCopiedActionMsg(false), 2500);
  };

  // Compare Two Snapshots Data with AI Strategic Insights Engine
  const comparisonViewData = useMemo(() => {
    const snapA = history.find(h => h.id === compareIdA);
    const snapB = history.find(h => h.id === compareIdB);
    if (!snapA || !snapB) return null;

    const aggA = aggregateRecords(snapA.rawRecords);
    const aggB = aggregateRecords(snapB.rawRecords);

    const asinMap = {};
    aggA.asins.forEach(a => {
      asinMap[a.asin] = {
        asin: a.asin,
        spendA: a.totalSpend,
        salesA: a.totalSales,
        acosA: a.acos,
        ordersA: a.totalOrders,
        spendB: 0,
        salesB: 0,
        acosB: 0,
        ordersB: 0,
        campaignsA: a.campaigns || [],
        campaignsB: []
      };
    });

    aggB.asins.forEach(b => {
      if (!asinMap[b.asin]) {
        asinMap[b.asin] = {
          asin: b.asin,
          spendA: 0,
          salesA: 0,
          acosA: 0,
          ordersA: 0,
          spendB: b.totalSpend,
          salesB: b.totalSales,
          acosB: b.acos,
          ordersB: b.totalOrders,
          campaignsA: [],
          campaignsB: b.campaigns || []
        };
      } else {
        asinMap[b.asin].spendB = b.totalSpend;
        asinMap[b.asin].salesB = b.totalSales;
        asinMap[b.asin].acosB = b.acos;
        asinMap[b.asin].ordersB = b.totalOrders;
        asinMap[b.asin].campaignsB = b.campaigns || [];
      }
    });

    const asinRows = Object.values(asinMap).map(row => {
      const spendDiff = row.spendB - row.spendA;
      const salesDiff = row.salesB - row.salesA;
      const acosDiff = row.acosB - row.acosA;
      const catItem = asinCatalog[row.asin];
      const aiReport = generateComparativeAIReport(row, catItem, globalStrategy);

      return {
        ...row,
        spendDiff,
        salesDiff,
        acosDiff,
        aiReport
      };
    }).sort((a, b) => {
      if (a.aiReport.priority !== b.aiReport.priority) {
        return a.aiReport.priority - b.aiReport.priority;
      }
      return (b.spendB + b.spendA) - (a.spendB + a.spendA);
    });

    const overallSummary = generateOverallComparativeSummary(
      aggA, 
      aggB, 
      asinRows.map(r => r.aiReport), 
      globalStrategy
    );

    return {
      snapA,
      snapB,
      aggA,
      aggB,
      asinRows,
      overallSummary
    };
  }, [history, compareIdA, compareIdB, asinCatalog, globalStrategy]);

  return (
    <div className="app-layout">
      {/* Sidebar Navigation */}
      <nav className="tabs-nav">
        <div className="app-brand">
          <BarChart2 size={24} color="var(--accent-color)" />
          <span>CM Analyzer</span>
        </div>
        
        <div className="tabs-list">
          <button 
            className={`tab-btn ${activeTab === 'upload' ? 'active' : ''}`}
            onClick={() => setActiveTab('upload')}
          >
            <FileSpreadsheet size={18} />
            Data Input
          </button>
          
          <button 
            className={`tab-btn ${activeTab === 'dashboard' ? 'active' : ''}`}
            onClick={() => setActiveTab('dashboard')}
          >
            <LayoutDashboard size={18} />
            Dashboard
          </button>

          <button 
            className={`tab-btn ${activeTab === 'ai_strategy' ? 'active' : ''}`}
            onClick={() => setActiveTab('ai_strategy')}
          >
            <Bot size={18} color="var(--accent-color)" />
            AI Strategy Engine
          </button>

          <button 
            className={`tab-btn ${activeTab === 'history' ? 'active' : ''}`}
            onClick={() => setActiveTab('history')}
          >
            <History size={18} />
            History ({history.length})
          </button>

          <button 
            className={`tab-btn ${activeTab === 'compare' ? 'active' : ''}`}
            onClick={() => setActiveTab('compare')}
          >
            <GitCompare size={18} />
            Compare Weeks
          </button>

          <button 
            className={`tab-btn ${activeTab === 'bulk_ops' ? 'active' : ''}`}
            onClick={() => setActiveTab('bulk_ops')}
          >
            <ListChecks size={18} />
            Bulk Operations
          </button>
          
          <button 
            className={`tab-btn ${activeTab === 'pesh_optimizer' ? 'active' : ''}`}
            onClick={() => setActiveTab('pesh_optimizer')}
            style={{ color: activeTab === 'pesh_optimizer' ? 'var(--purple)' : '' }}
          >
            <Target size={18} />
            PESH Optimizer
          </button>
          <button 
            className={`tab-btn ${activeTab === 'admin' ? 'active' : ''}`}
            onClick={() => setActiveTab('admin')}
            style={{ marginTop: 'auto', borderTop: '1px solid var(--border-color)', paddingTop: '1rem' }}
          >
            <ShieldCheck size={18} color={activeTab === 'admin' ? 'var(--accent-color)' : undefined} />
            Admin Panel
            {unreadFeedbackCount > 0 && (
              <span className="feedback-badge">{unreadFeedbackCount}</span>
            )}
          </button>
          <button 
            className="tab-btn"
            onClick={() => setIsFeedbackOpen(true)}
          >
            <MessageSquare size={18} />
            Feedback & Support
          </button>
        </div>
        
        {rawRecords.length > 0 && (
          <button className="nav-reset-btn" onClick={resetData}>
            <RefreshCw size={14} /> Clear Active Data
          </button>
        )}
      </nav>

      {/* Main Content Area */}
      <main className="tab-content">
        {/* TAB 1: UPLOAD */}
        {activeTab === 'upload' && (
          <div style={{display: 'flex', flexDirection: 'column', alignItems: 'center', margin: '1.5rem auto', maxWidth: '850px', width: '100%'}}>
            <div style={{display: 'flex', gap: '1.5rem', width: '100%', alignItems: 'stretch'}}>
              {/* Standard Report Upload */}
              <div 
                className={`drop-zone ${isDragOver ? 'drag-over' : ''} ${loading ? 'processing' : ''}`}
                style={{flex: 1, margin: 0, padding: '2.5rem 1.5rem'}}
                onDragOver={handleDragOver}
                onDragLeave={handleDragLeave}
                onDrop={handleDrop}
              >
                <FileSpreadsheet size={42} className="upload-icon" />
                <h2 style={{fontSize: '1.15rem', marginBottom: '0.5rem', color: 'var(--text-primary)'}}>Campaign Report</h2>
                <p className="upload-subtitle" style={{fontSize: '0.85rem', marginBottom: '1.5rem'}}>
                  For Dashboard, History, & ASIN Tracking.
                </p>
                
                <label className="upload-button">
                  {loading ? 'Processing...' : 'Upload Report'}
                  <input 
                    type="file" 
                    accept=".csv, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel" 
                    className="file-input" 
                    onChange={handleFileUpload}
                    disabled={loading}
                  />
                </label>
              </div>

              {/* Bulk File Upload */}
              <div 
                className={`drop-zone ${isDragOver ? 'drag-over' : ''} ${loading ? 'processing' : ''}`}
                style={{flex: 1, margin: 0, padding: '2.5rem 1.5rem', borderColor: 'rgba(168, 85, 247, 0.4)'}}
                onDragOver={handleDragOver}
                onDragLeave={handleDragLeave}
                onDrop={handleDrop}
              >
                <Target size={42} style={{color: 'var(--purple)', marginBottom: '1rem'}} />
                <h2 style={{fontSize: '1.15rem', marginBottom: '0.5rem', color: 'var(--text-primary)'}}>Bulk Operations File</h2>
                <p className="upload-subtitle" style={{fontSize: '0.85rem', marginBottom: '1.5rem'}}>
                  For the <b>PESH Optimizer</b> & Keyword actions.
                </p>
                
                <label className="upload-button" style={{backgroundColor: 'var(--purple)'}}>
                  {loading ? 'Processing...' : 'Upload Bulk File'}
                  <input 
                    type="file" 
                    accept=".csv, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel" 
                    className="file-input" 
                    onChange={handleFileUpload}
                    disabled={loading}
                  />
                </label>
              </div>
            </div>

            {rawRecords.length > 0 && (
              <div className="save-snapshot-box">
                <div className="save-box-title">
                  <Save size={16} color="var(--accent-color)" />
                  <span>Save this upload to History:</span>
                </div>
                <div className="save-input-row">
                  <input 
                    type="text" 
                    className="snapshot-input" 
                    value={snapshotNameInput} 
                    onChange={(e) => setSnapshotNameInput(e.target.value)}
                    placeholder="e.g. Week 32 - Aug 2026"
                  />
                  <button className="save-snapshot-btn" onClick={saveCurrentSnapshot}>
                    {savedSuccessMsg ? <CheckCircle2 size={16} color="#10b981"/> : <Plus size={16}/>}
                    {savedSuccessMsg ? 'Saved!' : 'Save Snapshot'}
                  </button>
                </div>
              </div>
            )}

            <div className="upload-tips">
              <div className="tip-header"><Sparkles size={16} color="var(--accent-color)"/> <strong>AI Strategy & Workflow:</strong></div>
              <p>The embedded AI Growth Agent automatically discovers ASINs, synchronizes product titles, and evaluates custom bid/harvesting strategies based on your targets.</p>
            </div>

            {fileName && (
              <p className="loaded-file-msg">Current active file: <strong>{fileName}</strong> ({rawRecords.length} rows)</p>
            )}
            
            {rawRecords.length > 0 && !loading && (
              <button className="go-to-dashboard-btn" onClick={() => setActiveTab('dashboard')}>
                View Dashboard →
              </button>
            )}
          </div>
        )}

        {/* TAB 2: DASHBOARD */}
        {activeTab === 'dashboard' && (
          <div className="dashboard-container">
            {!dashboardData ? (
              <div className="empty-state">
                <LayoutDashboard size={48} className="upload-icon" style={{opacity: 0.5}} />
                <h2>No active data loaded</h2>
                <p>Upload a campaign file or load a saved snapshot from your history.</p>
                <div style={{display: 'flex', gap: '1rem'}}>
                  <button className="upload-button" onClick={() => setActiveTab('upload')}>
                    Upload File
                  </button>
                  {history.length > 0 && (
                    <button className="go-to-dashboard-btn" onClick={() => setActiveTab('history')}>
                      Browse History ({history.length})
                    </button>
                  )}
                </div>
              </div>
            ) : (
              <>
                <div className="dashboard-header">
                  <div>
                    <h1 className="dashboard-title">ASIN Performance</h1>
                    {dateRangeInfo && (
                      <p className="date-span-subtitle">
                        <Calendar size={14} /> Report Span: {dateRangeInfo.minStr} – {dateRangeInfo.maxStr} ({dateRangeInfo.totalDays} days)
                      </p>
                    )}
                  </div>

                  <div style={{display: 'flex', alignItems: 'center', gap: '1rem'}}>
                    <button className="quick-ai-btn" onClick={() => setActiveTab('ai_strategy')}>
                      <Bot size={16} color="var(--accent-color)" />
                      <span>AI Strategy Suite</span>
                    </button>

                    <button className="quick-save-btn" onClick={saveCurrentSnapshot} title="Save to History">
                      {savedSuccessMsg ? <CheckCircle2 size={16} color="#10b981" /> : <Save size={16} />}
                      {savedSuccessMsg ? 'Saved' : 'Save Snapshot'}
                    </button>

                    {/* Time Range Filter Pills */}
                    <div className="period-filter-bar">
                      <button 
                        className={`period-pill ${selectedPeriod === 'all' ? 'active' : ''}`}
                        onClick={() => setSelectedPeriod('all')}
                      >
                        All Time
                      </button>
                      
                      <button 
                        className={`period-pill ${selectedPeriod === 'this_week' ? 'active' : ''} ${!hasDates ? 'disabled' : ''}`}
                        onClick={() => hasDates && setSelectedPeriod('this_week')}
                        disabled={!hasDates}
                      >
                        <Clock size={14} /> This Week
                      </button>

                      <button 
                        className={`period-pill ${selectedPeriod === 'last_week' ? 'active' : ''} ${!hasDates ? 'disabled' : ''}`}
                        onClick={() => hasDates && setSelectedPeriod('last_week')}
                        disabled={!hasDates}
                      >
                        Last Week
                      </button>

                      <button 
                        className={`period-pill ${selectedPeriod === 'last_30_days' ? 'active' : ''} ${!hasDates ? 'disabled' : ''}`}
                        onClick={() => hasDates && setSelectedPeriod('last_30_days')}
                        disabled={!hasDates}
                      >
                        Last 30 Days
                      </button>
                    </div>
                  </div>
                </div>

                {!hasDates && (
                  <div className="info-banner">
                    <Info size={18} />
                    <span>
                      Viewing summary data for full file. (For automated <strong>This Week / Last Week / 30 Days</strong> breakdown, export with <em>Daily</em> unit).
                    </span>
                  </div>
                )}

                {/* KPI Metrics Grid */}
                <div className="metrics-grid">
                  <div className="metric-card">
                    <div className="metric-header-row">
                      <span className="metric-title">Total Spend</span>
                      {dashboardData.previous && renderDelta(dashboardData.current.totalSpend, dashboardData.previous.totalSpend)}
                    </div>
                    <div className="metric-value">{formatCurrency(dashboardData.current.totalSpend)}</div>
                  </div>

                  <div className="metric-card">
                    <div className="metric-header-row">
                      <span className="metric-title">Total Sales</span>
                      {dashboardData.previous && renderDelta(dashboardData.current.totalSales, dashboardData.previous.totalSales)}
                    </div>
                    <div className="metric-value">{formatCurrency(dashboardData.current.totalSales)}</div>
                  </div>

                  <div className="metric-card">
                    <div className="metric-header-row">
                      <span className="metric-title">Total ACoS</span>
                      {dashboardData.previous && renderDelta(dashboardData.current.totalAcos, dashboardData.previous.totalAcos, true)}
                    </div>
                    <div className={`metric-value ${dashboardData.current.totalAcos > globalStrategy.targetAcos ? 'bad-acos' : 'good-acos'}`}>
                      {formatPercent(dashboardData.current.totalAcos)}
                    </div>
                  </div>

                  <div className="metric-card">
                    <div className="metric-header-row">
                      <span className="metric-title">Orders</span>
                      {dashboardData.previous && renderDelta(dashboardData.current.totalOrders, dashboardData.previous.totalOrders)}
                    </div>
                    <div className="metric-value">{formatNumber(dashboardData.current.totalOrders)}</div>
                  </div>

                  <div className="metric-card">
                    <div className="metric-header-row">
                      <span className="metric-title">ROAS</span>
                      {dashboardData.previous && renderDelta(dashboardData.current.totalRoas, dashboardData.previous.totalRoas)}
                    </div>
                    <div className={`metric-value ${dashboardData.current.totalRoas >= 3 ? 'good-acos' : dashboardData.current.totalRoas < 1.5 ? 'bad-acos' : ''}`}>
                      {dashboardData.current.totalRoas.toFixed(2)}x
                    </div>
                  </div>

                  <div className="metric-card">
                    <div className="metric-header-row">
                      <span className="metric-title">Avg CPC</span>
                      {dashboardData.previous && renderDelta(dashboardData.current.totalAvgCpc, dashboardData.previous.totalAvgCpc, true)}
                    </div>
                    <div className="metric-value">{formatCurrency(dashboardData.current.totalAvgCpc)}</div>
                  </div>

                  <div className="metric-card">
                    <div className="metric-header-row">
                      <span className="metric-title">CTR</span>
                      {dashboardData.previous && renderDelta(dashboardData.current.totalCtr, dashboardData.previous.totalCtr)}
                    </div>
                    <div className="metric-value">{dashboardData.current.totalCtr.toFixed(2)}%</div>
                  </div>

                  <div className="metric-card">
                    <div className="metric-header-row">
                      <span className="metric-title">CVR</span>
                      {dashboardData.previous && renderDelta(dashboardData.current.totalCvr, dashboardData.previous.totalCvr)}
                    </div>
                    <div className="metric-value">{dashboardData.current.totalCvr.toFixed(2)}%</div>
                  </div>
                </div>

                {/* EASY WIN CAMPAIGNS SECTION */}
                {dashboardData.current.easyWin?.asins?.length > 0 && (() => {
                  const ew = dashboardData.current.easyWin;
                  return (
                    <div className="ew-section">
                      <div className="ew-section-header">
                        <div className="ew-section-title-row">
                          <span className="ew-icon">⚡</span>
                          <h2 className="ew-section-title">Easy Win Campaigns</h2>
                          <span className="ew-campaign-count">{ew.asins.reduce((s, g) => s + g.campaigns.length, 0)} campaigns across {ew.asins.length} ASINs</span>
                        </div>
                        {/* Overall EW KPIs */}
                        <div className="ew-overall-kpis">
                          <div className="ew-kpi">
                            <span className="ew-kpi-label">Total Spend</span>
                            <span className="ew-kpi-value">{formatCurrency(ew.totalSpend)}</span>
                          </div>
                          <div className="ew-kpi">
                            <span className="ew-kpi-label">Total Sales</span>
                            <span className="ew-kpi-value">{formatCurrency(ew.totalSales)}</span>
                          </div>
                          <div className="ew-kpi">
                            <span className="ew-kpi-label">Overall ACoS</span>
                            <span className={`ew-kpi-value ${ew.isBleeding ? 'bad-acos' : ew.totalAcos > globalStrategy.targetAcos ? 'bad-acos' : 'good-acos'}`}>
                              {ew.isBleeding ? 'N/A ⚠️' : formatPercent(ew.totalAcos)}
                            </span>
                          </div>
                        </div>
                      </div>

                      {/* ASIN Subcategories */}
                      <div className="ew-asin-list">
                        {ew.asins.map((ewAsin) => {
                          const catItem = asinCatalog[ewAsin.asin];
                          const productName = catItem?.productName && catItem.productName !== `Product ${ewAsin.asin}`
                            ? catItem.productName
                            : null;
                          const ewAsinKey = `ew_${ewAsin.asin}`;
                          return (
                            <div key={ewAsinKey} className="ew-asin-item">
                              <div className="ew-asin-header" onClick={() => toggleAsin(ewAsinKey)}>
                                <div className="ew-asin-left">
                                  <ChevronDown size={18} className={`chevron-icon ${expandedAsin === ewAsinKey ? 'open' : ''}`} />
                                  <div>
                                    <div className="ew-asin-name-row">
                                      <span className="asin-name">{ewAsin.asin}</span>
                                      {productName && <span className="asin-product-title" title={productName}>{productName}</span>}
                                    </div>
                                    <span style={{fontSize:'0.75rem',color:'var(--text-secondary)'}}>
                                      {ewAsin.campaigns.length} Easy Win campaign{ewAsin.campaigns.length !== 1 ? 's' : ''}
                                    </span>
                                  </div>
                                </div>
                                <div className="ew-asin-kpis">
                                  <div className="ew-asin-kpi">
                                    <span className="ew-kpi-label">Spend</span>
                                    <span className="ew-kpi-value">{formatCurrency(ewAsin.spend)}</span>
                                  </div>
                                  <div className="ew-asin-kpi">
                                    <span className="ew-kpi-label">Sales</span>
                                    <span className="ew-kpi-value">{formatCurrency(ewAsin.sales)}</span>
                                  </div>
                                  <div className="ew-asin-kpi">
                                    <span className="ew-kpi-label">ACoS</span>
                                    <span className={`ew-kpi-value ${ewAsin.isBleeding ? 'bad-acos' : ewAsin.acos > globalStrategy.targetAcos ? 'bad-acos' : 'good-acos'}`}>
                                      {ewAsin.isBleeding ? 'N/A ⚠️' : formatPercent(ewAsin.acos)}
                                    </span>
                                  </div>
                                </div>
                              </div>

                              {expandedAsin === ewAsinKey && (
                                <div className="campaigns-container" style={{marginTop: '0.5rem', paddingLeft: '2rem'}}>
                                  <table className="campaign-table">
                                    <thead>
                                      <tr>
                                        <th>Campaign Name</th>
                                        <th className="text-right">Spend</th>
                                        <th className="text-right">Sales</th>
                                        <th className="text-right">ACoS</th>
                                        <th className="text-right">Orders</th>
                                      </tr>
                                    </thead>
                                    <tbody>
                                      {ewAsin.campaigns.map((camp, idx) => (
                                        <tr key={idx}>
                                          <td className="campaign-name-cell" title={camp.name}>{camp.name}</td>
                                          <td className="text-right">{formatCurrency(camp.spend)}</td>
                                          <td className="text-right">{formatCurrency(camp.sales)}</td>
                                          <td className={`text-right ${camp.isBleeding ? 'bad-acos' : camp.acos > globalStrategy.targetAcos ? 'bad-acos' : 'good-acos'}`}>
                                            {camp.isBleeding ? 'N/A ⚠️' : formatPercent(camp.acos)}
                                          </td>
                                          <td className="text-right">{formatNumber(camp.orders)}</td>
                                        </tr>
                                      ))}
                                    </tbody>
                                  </table>
                                </div>
                              )}
                            </div>
                          );
                        })}
                      </div>
                    </div>
                  );
                })()}

                {/* ASIN Breakdown Accordion List */}
                <div className="asin-list">
                  {dashboardData.current.asins.length === 0 ? (
                    <div className="no-period-data">
                      <p>No campaign activity recorded during this specific time window.</p>
                      <button className="period-pill active" onClick={() => setSelectedPeriod('all')}>
                        Show All Time Data
                      </button>
                    </div>
                  ) : (
                    dashboardData.current.asins.map((asinGroup) => {
                      const catItem = asinCatalog[asinGroup.asin];
                      const title = catItem?.productName || asinGroup.asin;
                      const isCustom = catItem?.strategyMode && catItem.strategyMode !== 'global';

                      return (
                        <div key={asinGroup.asin} className="asin-item">
                          <div className="asin-header" onClick={() => toggleAsin(asinGroup.asin)}>
                            <div className="asin-title-section">
                              <ChevronDown 
                                size={24} 
                                className={`chevron-icon ${expandedAsin === asinGroup.asin ? 'open' : ''}`} 
                              />
                              <div className="asin-identity-block">
                                <div className="asin-title-row">
                                  <span className="asin-name">{asinGroup.asin}</span>
                                  {catItem?.productName && catItem.productName !== `Product ${asinGroup.asin}` && (
                                    <span className="asin-product-title" title={catItem.productName}>
                                      {catItem.productName}
                                    </span>
                                  )}
                                  <button 
                                    className="edit-asin-meta-btn"
                                    onClick={(e) => openCatalogEditor(asinGroup.asin, e)}
                                    title="Edit Product Details & Custom Strategy"
                                  >
                                    <Edit2 size={13} />
                                  </button>
                                </div>

                                <div className="asin-sub-tags">
                                  {asinGroup.perfCategory === 'top_performer' && (
                                    <span className="perf-badge perf-top">🏆 Top Performer</span>
                                  )}
                                  {asinGroup.perfCategory === 'underperformer' && (
                                    <span className="perf-badge perf-under">🚨 Underperformer</span>
                                  )}
                                  {asinGroup.isBleeding && (
                                    <span className="perf-badge perf-bleed">💸 Bleeding</span>
                                  )}
                                  <span className="campaign-count-tag">
                                    {asinGroup.campaigns.length} campaigns
                                  </span>
                                  {catItem?.category && catItem.category !== 'General E-Commerce' && (
                                    <span className="asin-category-tag">
                                      {catItem.category}
                                    </span>
                                  )}
                                  {catItem?.autoDetected && (
                                    <span className="ai-detected-badge" title="Product name & category auto-detected by AI">
                                      <Sparkles size={11} /> AI Title
                                    </span>
                                  )}
                                  {isCustom && (
                                    <span className="custom-strategy-tag">
                                      <Target size={12} /> Custom: {catItem.strategyMode}
                                    </span>
                                  )}
                                </div>
                              </div>
                            </div>

                            <div className="asin-header-right-side">
                              {/* Direct AI Audit button */}
                              <button 
                                className="asin-ai-action-btn"
                                onClick={(e) => openAIModalForAsin(asinGroup, e)}
                                title="Run Autonomous AI Optimization Plan"
                              >
                                <Zap size={14} /> AI Actions
                              </button>

                              <div className="asin-summary-metrics">
                                <div className="summary-metric">
                                  <span className="summary-label">Spend</span>
                                  <span className="summary-value">{formatCurrency(asinGroup.totalSpend)}</span>
                                </div>
                                <div className="summary-metric">
                                  <span className="summary-label">Sales</span>
                                  <span className="summary-value">{formatCurrency(asinGroup.totalSales)}</span>
                                </div>
                                <div className="summary-metric">
                                  <span className="summary-label">ACoS</span>
                                  <span className={`summary-value ${asinGroup.isBleeding ? 'bad-acos' : asinGroup.acos > (catItem?.targetAcos || globalStrategy.targetAcos) ? 'bad-acos' : 'good-acos'}`}>
                                    {asinGroup.isBleeding ? 'N/A ⚠️' : formatPercent(asinGroup.acos)}
                                  </span>
                                </div>
                              </div>
                            </div>
                          </div>
                          
                          {expandedAsin === asinGroup.asin && (
                            <div className="campaigns-container">
                              <table className="campaign-table">
                                <thead>
                                  <tr>
                                    <th>Campaign Name</th>
                                    <th className="text-right">Budget</th>
                                    <th className="text-right">Spend</th>
                                    <th className="text-right">Sales</th>
                                    <th className="text-right">ACoS</th>
                                    <th className="text-right">Orders</th>
                                    <th className="text-right">Impr</th>
                                    <th className="text-right">Clicks</th>
                                    <th className="text-right">CPC</th>
                                    <th>Signals & Suggestions</th>
                                  </tr>
                                </thead>
                                <tbody>
                                  {asinGroup.campaigns.map((camp, idx) => (
                                    <tr key={idx}>
                                      <td className="campaign-name-cell" title={camp.name}>{camp.name}</td>
                                      <td className="text-right">{formatCurrency(camp.budget)}</td>
                                      <td className="text-right">{formatCurrency(camp.spend)}</td>
                                      <td className="text-right">{formatCurrency(camp.sales)}</td>
                                      <td className={`text-right ${camp.isBleeding ? 'bad-acos' : camp.acos > globalStrategy.targetAcos ? 'bad-acos' : 'good-acos'}`}>
                                        {camp.isBleeding ? 'N/A ⚠️' : formatPercent(camp.acos)}
                                      </td>
                                      <td className="text-right">{formatNumber(camp.orders)}</td>
                                      <td className="text-right">{formatNumber(camp.impressions)}</td>
                                      <td className="text-right">{formatNumber(camp.clicks)}</td>
                                      <td className="text-right">{formatCurrency(camp.cpc)}</td>
                                      <td className="signals-cell">
                                        {camp.signals.length > 0 ? (
                                          camp.signals.map((sig, i) => (
                                            <div key={i} className={`signal-badge signal-${sig.type}`} title={sig.msg}>
                                              {sig.icon}
                                              <span>{sig.msg}</span>
                                            </div>
                                          ))
                                        ) : (
                                          <span className="no-signals">-</span>
                                        )}
                                      </td>
                                    </tr>
                                  ))}
                                </tbody>
                              </table>
                            </div>
                          )}
                        </div>
                      );
                    })
                  )}
                </div>
              </>
            )}
          </div>
        )}

        {/* TAB 3: AI STRATEGY ENGINE */}
        {activeTab === 'ai_strategy' && (
          <div className="dashboard-container">
            <div className="dashboard-header">
              <div>
                <h1 className="dashboard-title">
                  <Bot size={28} color="var(--accent-color)" style={{verticalAlign: 'middle', marginRight: '0.5rem'}} />
                  Autonomous AI Strategy Engine
                </h1>
                <p className="date-span-subtitle">
                  Data-driven PPC decisions, automated keyword harvesting, bid optimizations, and ASIN-by-ASIN custom strategy directives.
                </p>
              </div>
            </div>

            {/* Global Strategy Configuration Card */}
            <div className="strategy-settings-card">
              <div className="settings-card-header">
                <div style={{display: 'flex', alignItems: 'center', gap: '0.6rem'}}>
                  <Sliders size={20} color="var(--accent-color)" />
                  <div>
                    <h3 style={{margin: 0, fontSize: '1.1rem'}}>Global Portfolio Strategy Settings</h3>
                    <span style={{fontSize: '0.8rem', color: 'var(--text-secondary)'}}>
                      Configure master algorithmic targets, default objectives, and spending cutoffs for the portfolio
                    </span>
                  </div>
                </div>
                <div style={{display: 'flex', alignItems: 'center', gap: '0.6rem'}}>
                  {strategySavedToast ? (
                    <span className="strategy-status-indicator toast-saved">
                      <Check size={14} /> Saved to Memory
                    </span>
                  ) : (
                    <span className="strategy-status-indicator">Active Heuristic Engine</span>
                  )}
                </div>
              </div>

              {/* 1-Click Strategy Presets Bar */}
              <div className="strategy-presets-bar">
                <span className="presets-label">
                  <Sparkles size={14} color="var(--accent-color)" /> Quick Strategy Presets:
                </span>
                <div className="preset-buttons-group">
                  {STRATEGY_PRESETS.map(preset => {
                    const isActive = globalStrategy.defaultObjective === preset.config.defaultObjective && 
                                     Number(globalStrategy.targetAcos) === preset.config.targetAcos &&
                                     Number(globalStrategy.bleedSpendThreshold) === preset.config.bleedSpendThreshold;
                    return (
                      <button
                        key={preset.id}
                        type="button"
                        className={`preset-btn ${isActive ? 'active-preset' : ''}`}
                        onClick={() => applyStrategyPreset(preset)}
                        title={preset.desc}
                      >
                        <span>{preset.badge}</span>
                      </button>
                    );
                  })}
                </div>
              </div>

              {/* Interactive Settings Form Grid with Steppers */}
              <div className="settings-form-grid">
                <div className="form-group">
                  <label>Default Strategy Objective</label>
                  <select 
                    value={globalStrategy.defaultObjective} 
                    onChange={(e) => handleGlobalStrategyChange('defaultObjective', e.target.value)}
                    className="form-control"
                  >
                    <option value="growth">Growth / Scale (Harvesting & Ranking)</option>
                    <option value="launch">Launch / Cold Start (Aggressive Discovery)</option>
                    <option value="profitability">Profitability & Margin Protection (Strict ACoS)</option>
                    <option value="defense">Defense (Protect Brand & PDPs)</option>
                  </select>
                  <span className="input-helper">Master objective applied to all ASINs without custom override</span>
                </div>

                <div className="form-group">
                  <label>Portfolio Target ACoS (%)</label>
                  <div className="stepper-input-wrapper">
                    <button 
                      type="button" 
                      className="stepper-btn"
                      onClick={() => adjustNumericStrategy('targetAcos', -1, 1, 100)}
                      title="Decrease by 1%"
                    >
                      <Minus size={14} />
                    </button>
                    <input 
                      type="number" 
                      step="1"
                      min="1"
                      max="100"
                      value={globalStrategy.targetAcos ?? ''} 
                      onChange={(e) => handleGlobalStrategyChange('targetAcos', e.target.value)}
                      onBlur={() => handleGlobalStrategyBlur('targetAcos', 30)}
                      className="form-control stepper-input"
                    />
                    <button 
                      type="button" 
                      className="stepper-btn"
                      onClick={() => adjustNumericStrategy('targetAcos', 1, 1, 100)}
                      title="Increase by 1%"
                    >
                      <Plus size={14} />
                    </button>
                  </div>
                  <span className="input-helper">Target ad spend to sales efficiency benchmark</span>
                </div>

                <div className="form-group">
                  <label>Portfolio Target TACoS (%)</label>
                  <div className="stepper-input-wrapper">
                    <button 
                      type="button" 
                      className="stepper-btn"
                      onClick={() => adjustNumericStrategy('targetTacos', -1, 1, 100)}
                      title="Decrease by 1%"
                    >
                      <Minus size={14} />
                    </button>
                    <input 
                      type="number" 
                      step="1"
                      min="1"
                      max="100"
                      value={globalStrategy.targetTacos ?? ''} 
                      onChange={(e) => handleGlobalStrategyChange('targetTacos', e.target.value)}
                      onBlur={() => handleGlobalStrategyBlur('targetTacos', 12)}
                      className="form-control stepper-input"
                    />
                    <button 
                      type="button" 
                      className="stepper-btn"
                      onClick={() => adjustNumericStrategy('targetTacos', 1, 1, 100)}
                      title="Increase by 1%"
                    >
                      <Plus size={14} />
                    </button>
                  </div>
                  <span className="input-helper">Target total ad spend vs total product revenue</span>
                </div>

                <div className="form-group">
                  <label>Bleeding Spend Cutoff ($ with 0 orders)</label>
                  <div className="stepper-input-wrapper">
                    <button 
                      type="button" 
                      className="stepper-btn"
                      onClick={() => adjustNumericStrategy('bleedSpendThreshold', -5, 1, 500)}
                      title="Decrease by $5"
                    >
                      <Minus size={14} />
                    </button>
                    <input 
                      type="number" 
                      step="1"
                      min="1"
                      max="500"
                      value={globalStrategy.bleedSpendThreshold ?? ''} 
                      onChange={(e) => handleGlobalStrategyChange('bleedSpendThreshold', e.target.value)}
                      onBlur={() => handleGlobalStrategyBlur('bleedSpendThreshold', 15)}
                      className="form-control stepper-input"
                    />
                    <button 
                      type="button" 
                      className="stepper-btn"
                      onClick={() => adjustNumericStrategy('bleedSpendThreshold', 5, 1, 500)}
                      title="Increase by $5"
                    >
                      <Plus size={14} />
                    </button>
                  </div>
                  <span className="input-helper">Spend threshold triggering immediate bid cuts / pauses</span>
                </div>
              </div>

              {/* Action Toolbar for Settings */}
              <div className="strategy-actions-toolbar">
                <div className="left-actions">
                  <button 
                    type="button"
                    className="strategy-tool-btn apply-all-btn"
                    onClick={applyTargetAcosToAllCatalog}
                    title="Push current Target ACoS to all ASINs in catalog"
                  >
                    <Zap size={14} />
                    <span>Apply Target ACoS ({globalStrategy.targetAcos}%) to All Products</span>
                  </button>
                </div>
                <div className="right-actions">
                  <button 
                    type="button"
                    className="strategy-tool-btn reset-btn"
                    onClick={resetGlobalStrategyToDefaults}
                    title="Reset all settings to default baseline"
                  >
                    <RotateCcw size={14} />
                    <span>Reset Defaults</span>
                  </button>
                  <button 
                    type="button"
                    className="strategy-tool-btn save-btn"
                    onClick={() => {
                      setStrategySavedToast(true);
                      setTimeout(() => setStrategySavedToast(false), 2500);
                    }}
                  >
                    <Save size={14} />
                    <span>Save & Apply Settings</span>
                  </button>
                </div>
              </div>
            </div>

            {/* ASIN Strategy & Autonomous Action Table */}
            <div className="ai-matrix-container">
              <div className="matrix-header" style={{display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '1rem'}}>
                <div>
                  <h3>ASIN Strategy Matrix & Optimization Plans</h3>
                  <span style={{fontSize: '0.85rem', color: 'var(--text-secondary)'}}>
                    Click "⚡ Generate AI Plan" on any ASIN for full step-by-step optimization directives.
                  </span>
                </div>
                {rawRecords.length > 0 && (
                  <button 
                    className="batch-ai-detect-btn" 
                    onClick={redetectAllAsinTitles}
                    title="Re-run AI extraction across all campaigns to detect and update product names"
                  >
                    <Sparkles size={14} />
                    <span>Auto-Detect All ASIN Names (AI)</span>
                  </button>
                )}
              </div>

              {!dashboardData ? (
                <p style={{padding: '2rem', textAlign: 'center', color: 'var(--text-secondary)'}}>
                  Please upload a dataset to generate real-time AI optimization plans.
                </p>
              ) : (
                <table className="campaign-table">
                  <thead>
                    <tr>
                      <th>ASIN & Product Title</th>
                      <th>Category & Brand</th>
                      <th>Strategy Mode</th>
                      <th className="text-right">Target ACoS</th>
                      <th className="text-right">Actual ACoS</th>
                      <th className="text-right">Spend</th>
                      <th className="text-right">Sales</th>
                      <th>AI Optimization Action</th>
                    </tr>
                  </thead>
                  <tbody>
                    {dashboardData.current.asins.map(group => {
                      const catItem = asinCatalog[group.asin];
                      const title = catItem?.productName || group.asin;
                      const brand = catItem?.brand || 'Store Brand';
                      const category = catItem?.category || 'Amazon Catalog';
                      const isCustom = catItem?.strategyMode && catItem.strategyMode !== 'global';
                      const mode = isCustom ? catItem.strategyMode : globalStrategy.defaultObjective;
                      const targetAcos = catItem?.targetAcos || globalStrategy.targetAcos;

                      return (
                        <tr key={group.asin}>
                          <td>
                            <div style={{fontWeight: 600, color: 'var(--accent-color)'}}>{group.asin}</div>
                            <div style={{fontSize: '0.85rem', color: 'var(--text-primary)', fontWeight: 500, maxWidth: '280px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap'}}>
                              {title}
                            </div>
                          </td>
                          <td>
                            <div style={{fontSize: '0.8rem', color: 'var(--text-primary)'}}>{category}</div>
                            <div style={{fontSize: '0.75rem', color: 'var(--text-secondary)'}}>{brand}</div>
                          </td>
                          <td>
                            <span className={`strategy-pill ${isCustom ? 'custom-pill' : 'global-pill'}`}>
                              {isCustom ? `[Override] ${mode}` : `[Global] ${mode}`}
                            </span>
                          </td>
                          <td className="text-right">{formatPercent(targetAcos)}</td>
                          <td className={`text-right ${group.acos > targetAcos ? 'bad-acos' : 'good-acos'}`}>
                            {formatPercent(group.acos)}
                          </td>
                          <td className="text-right">{formatCurrency(group.totalSpend)}</td>
                          <td className="text-right">{formatCurrency(group.totalSales)}</td>
                          <td>
                            <div style={{display: 'flex', gap: '0.5rem'}}>
                              <button 
                                className="ai-plan-btn"
                                onClick={() => openAIModalForAsin(group)}
                              >
                                <Zap size={14} /> Generate AI Plan
                              </button>
                              <button 
                                className="edit-mini-btn"
                                onClick={() => openCatalogEditor(group.asin)}
                                title="Configure ASIN Custom Strategy"
                              >
                                <Edit2 size={13} />
                              </button>
                            </div>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              )}
            </div>
          </div>
        )}

        {/* TAB 4: HISTORY */}
        {activeTab === 'history' && (
          <div className="dashboard-container">
            <div className="dashboard-header">
              <div>
                <h1 className="dashboard-title">Saved Weekly History</h1>
                <p className="date-span-subtitle">Snapshots saved from your past uploads. Restores directly to the dashboard anytime.</p>
              </div>
            </div>

            {history.length === 0 ? (
              <div className="empty-state">
                <History size={48} className="upload-icon" style={{opacity: 0.5}} />
                <h2>No saved history yet</h2>
                <p>When you upload a weekly report, click <strong>"Save Snapshot"</strong> to preserve it here.</p>
                <button className="upload-button" onClick={() => setActiveTab('upload')}>
                  Go to Upload
                </button>
              </div>
            ) : (
              <div className="history-grid">
                {history.map((snap) => (
                  <div key={snap.id} className="history-card" onClick={() => loadSnapshot(snap)}>
                    <div className="history-card-header">
                      <div>
                        <h3 className="history-title">{snap.name}</h3>
                        <span className="history-date">Saved on {new Date(snap.savedAt).toLocaleDateString()} • {snap.fileName}</span>
                      </div>
                      <button className="delete-snap-btn" onClick={(e) => deleteSnapshot(snap.id, e)} title="Delete snapshot">
                        <Trash2 size={16} />
                      </button>
                    </div>

                    <div className="history-metrics-row">
                      <div className="hist-metric">
                        <span className="hist-label">Spend</span>
                        <span className="hist-val">{formatCurrency(snap.summary.totalSpend)}</span>
                      </div>
                      <div className="hist-metric">
                        <span className="hist-label">Sales</span>
                        <span className="hist-val">{formatCurrency(snap.summary.totalSales)}</span>
                      </div>
                      <div className="hist-metric">
                        <span className="hist-label">ACoS</span>
                        <span className={`hist-val ${snap.summary.totalAcos > globalStrategy.targetAcos ? 'bad-acos' : 'good-acos'}`}>
                          {formatPercent(snap.summary.totalAcos)}
                        </span>
                      </div>
                      <div className="hist-metric">
                        <span className="hist-label">ASINs</span>
                        <span className="hist-val">{snap.summary.asinCount}</span>
                      </div>
                    </div>

                    <div className="history-card-footer">
                      <button className="load-snap-btn">
                        <FolderOpen size={14} /> Open in Dashboard
                      </button>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>
        )}

        {/* TAB 5: COMPARE WEEKS */}
        {activeTab === 'compare' && (
          <div className="dashboard-container">
            <div className="dashboard-header">
              <div>
                <h1 className="dashboard-title">Compare Snapshots & Weeks</h1>
                <p className="date-span-subtitle">Select any two saved snapshots to see ASIN-by-ASIN performance evolution side-by-side.</p>
              </div>
            </div>

            {history.length < 2 ? (
              <div className="empty-state">
                <GitCompare size={48} className="upload-icon" style={{opacity: 0.5}} />
                <h2>Need at least 2 saved snapshots</h2>
                <p>Save reports from multiple weeks in the Data Input tab to compare them side by side.</p>
                <button className="upload-button" onClick={() => setActiveTab('upload')}>
                  Upload & Save Data
                </button>
              </div>
            ) : (
              <>
                <div className="compare-selectors-bar">
                  <div className="selector-group">
                    <label>Base Period (Week A):</label>
                    <select value={compareIdA} onChange={(e) => setCompareIdA(e.target.value)} className="compare-select">
                      <option value="">-- Select Period A --</option>
                      {history.map(h => (
                        <option key={h.id} value={h.id}>{h.name} ({new Date(h.savedAt).toLocaleDateString()})</option>
                      ))}
                    </select>
                  </div>

                  <div className="compare-vs-badge">VS</div>

                  <div className="selector-group">
                    <label>Comparison Period (Week B):</label>
                    <select value={compareIdB} onChange={(e) => setCompareIdB(e.target.value)} className="compare-select">
                      <option value="">-- Select Period B --</option>
                      {history.map(h => (
                        <option key={h.id} value={h.id}>{h.name} ({new Date(h.savedAt).toLocaleDateString()})</option>
                      ))}
                    </select>
                  </div>
                </div>

                {comparisonViewData ? (
                  <div className="compare-results">
                    {/* TOP METRICS GRID */}
                    <div className="metrics-grid">
                      <div className="metric-card">
                        <div className="metric-title">Spend Change</div>
                        <div className="metric-value">
                          {formatCurrency(comparisonViewData.aggB.totalSpend)}
                          <span style={{fontSize: '0.9rem', color: 'var(--text-secondary)', marginLeft: '0.5rem'}}>
                            (was {formatCurrency(comparisonViewData.aggA.totalSpend)})
                          </span>
                        </div>
                        {renderDelta(comparisonViewData.aggB.totalSpend, comparisonViewData.aggA.totalSpend)}
                      </div>

                      <div className="metric-card">
                        <div className="metric-title">Sales Change</div>
                        <div className="metric-value">
                          {formatCurrency(comparisonViewData.aggB.totalSales)}
                          <span style={{fontSize: '0.9rem', color: 'var(--text-secondary)', marginLeft: '0.5rem'}}>
                            (was {formatCurrency(comparisonViewData.aggA.totalSales)})
                          </span>
                        </div>
                        {renderDelta(comparisonViewData.aggB.totalSales, comparisonViewData.aggA.totalSales)}
                      </div>

                      <div className="metric-card">
                        <div className="metric-title">ACoS Change</div>
                        <div className="metric-value">
                          {formatPercent(comparisonViewData.aggB.totalAcos)}
                          <span style={{fontSize: '0.9rem', color: 'var(--text-secondary)', marginLeft: '0.5rem'}}>
                            (was {formatPercent(comparisonViewData.aggA.totalAcos)})
                          </span>
                        </div>
                        {renderDelta(comparisonViewData.aggB.totalAcos, comparisonViewData.aggA.totalAcos, true)}
                      </div>
                    </div>

                    {/* AI STRATEGIC COMPARATIVE BRIEF */}
                    {comparisonViewData.overallSummary && (
                      <div className="compare-ai-summary-card">
                        <div className="compare-ai-summary-header">
                          <div style={{display: 'flex', alignItems: 'center', gap: '0.6rem'}}>
                            <Sparkles size={20} color="var(--accent-color)" />
                            <h3 style={{margin: 0, fontSize: '1.05rem', fontWeight: 700}}>
                              AI Strategic Comparative Brief (WoW Analysis)
                            </h3>
                          </div>
                          <span className={`compare-status-badge status-${comparisonViewData.overallSummary.trendType}`}>
                            {comparisonViewData.overallSummary.trendType === 'success' && <Rocket size={14} />}
                            {comparisonViewData.overallSummary.trendType === 'danger' && <AlertTriangle size={14} />}
                            {comparisonViewData.overallSummary.trendType === 'info' && <TrendingUp size={14} />}
                            {comparisonViewData.overallSummary.trendType === 'neutral' && <CheckCircle2 size={14} />}
                            {comparisonViewData.overallSummary.trendHeadline}
                          </span>
                        </div>

                        <p className="compare-ai-narrative">
                          {comparisonViewData.overallSummary.executiveSummary}
                        </p>

                        <div className="compare-ai-pills-row">
                          <button 
                            className={`compare-filter-pill ${compareFilter === 'all' ? 'active' : ''}`}
                            onClick={() => setCompareFilter('all')}
                          >
                            All ASINs ({comparisonViewData.asinRows.length})
                          </button>

                          {comparisonViewData.overallSummary.actionCount > 0 && (
                            <button 
                              className={`compare-filter-pill pill-danger ${compareFilter === 'action_needed' ? 'active' : ''}`}
                              onClick={() => setCompareFilter('action_needed')}
                            >
                              <AlertCircle size={13} />
                              🚨 Action Required ({comparisonViewData.overallSummary.actionCount})
                            </button>
                          )}

                          {comparisonViewData.overallSummary.winnerCount > 0 && (
                            <button 
                              className={`compare-filter-pill pill-success ${compareFilter === 'winners' ? 'active' : ''}`}
                              onClick={() => setCompareFilter('winners')}
                            >
                              <Rocket size={13} />
                              🚀 Scale Winners ({comparisonViewData.overallSummary.winnerCount})
                            </button>
                          )}

                          {comparisonViewData.overallSummary.marginCount > 0 && (
                            <button 
                              className={`compare-filter-pill pill-info ${compareFilter === 'margin_improved' ? 'active' : ''}`}
                              onClick={() => setCompareFilter('margin_improved')}
                            >
                              <ShieldCheck size={13} />
                              🛡️ Margin Improved ({comparisonViewData.overallSummary.marginCount})
                            </button>
                          )}
                        </div>
                      </div>
                    )}

                    {/* COMPARISON TABLE */}
                    <div className="comparison-table-container">
                      <table className="campaign-table">
                        <thead>
                          <tr>
                            <th>ASIN & Product</th>
                            <th className="text-right">Spend (A)</th>
                            <th className="text-right">Spend (B)</th>
                            <th className="text-right">Δ Spend</th>
                            <th className="text-right">Sales (A)</th>
                            <th className="text-right">Sales (B)</th>
                            <th className="text-right">Δ Sales</th>
                            <th className="text-right">ACoS (A)</th>
                            <th className="text-right">ACoS (B)</th>
                            <th className="text-right">Δ ACoS</th>
                            <th>AI Suggestion</th>
                            <th style={{minWidth: '170px'}}>📝 Notes</th>
                            <th className="text-center">Action</th>
                          </tr>
                        </thead>
                        <tbody>
                          {comparisonViewData.asinRows
                            .filter(row => {
                              if (compareFilter === 'all') return true;
                              return row.aiReport.categoryKey === compareFilter;
                            })
                            .map((row) => {
                              const report = row.aiReport;
                              const catItem = asinCatalog[row.asin];
                              const displayTitle = catItem?.productName || (row.asin === 'Unknown' ? 'Non-ASIN Grouped Campaigns' : `Product (${row.asin})`);

                              return (
                                <tr key={row.asin}>
                                  <td>
                                    <div className="compare-product-cell">
                                      <div className="compare-asin-title" title={displayTitle}>
                                        {displayTitle}
                                      </div>
                                      <div className="compare-asin-meta">
                                        <span className="compare-asin-code">{row.asin}</span>
                                        <span className={`strategy-pill mode-${report.mode}`}>
                                          {report.mode.toUpperCase()}
                                        </span>
                                      </div>
                                    </div>
                                  </td>
                                  <td className="text-right">{formatCurrency(row.spendA)}</td>
                                  <td className="text-right">{formatCurrency(row.spendB)}</td>
                                  <td className="text-right">{renderDelta(row.spendB, row.spendA)}</td>
                                  <td className="text-right">{formatCurrency(row.salesA)}</td>
                                  <td className="text-right">{formatCurrency(row.salesB)}</td>
                                  <td className="text-right">{renderDelta(row.salesB, row.salesA)}</td>
                                  <td className="text-right">{formatPercent(row.acosA)}</td>
                                  <td className="text-right">{formatPercent(row.acosB)}</td>
                                  <td className="text-right">{renderDelta(row.acosB, row.acosA, true)}</td>
                                  <td>
                                    <div className="compare-ai-suggestion-cell">
                                      <span className={`compare-ai-badge badge-${report.badge.type}`}>
                                        {report.badge.label}
                                      </span>
                                      <div className="compare-ai-text" title={report.shortSuggestion}>
                                        {report.shortSuggestion}
                                      </div>
                                    </div>
                                  </td>
                                  <td>
                                    <input
                                      className="compare-note-input"
                                      type="text"
                                      placeholder="Add action note..."
                                      value={compareNotes[row.asin] || ''}
                                      onChange={(e) => setCompareNotes(prev => ({ ...prev, [row.asin]: e.target.value }))}
                                    />
                                  </td>
                                  <td className="text-center">
                                    <button
                                      className="compare-ai-action-btn"
                                      onClick={() => setSelectedCompareAsinForAI({
                                        row,
                                        report,
                                        snapA: comparisonViewData.snapA,
                                        snapB: comparisonViewData.snapB
                                      })}
                                      title="View Detailed Comparative AI Action Plan"
                                    >
                                      <Bot size={13} />
                                      <span>AI Plan</span>
                                    </button>
                                  </td>
                                </tr>
                              );
                            })}
                        </tbody>
                      </table>
                    </div>

                    {/* EXPORT BAR */}
                    <div className="compare-export-bar">
                      <div className="compare-export-left">
                        <ExternalLink size={16} color="var(--accent-color)" />
                        <span className="compare-export-label">Export Comparison</span>
                        <label className="compare-export-toggle">
                          <input
                            type="checkbox"
                            checked={exportAllRows}
                            onChange={(e) => setExportAllRows(e.target.checked)}
                          />
                          Include all rows (not just noted)
                        </label>
                        <span style={{color: 'var(--text-secondary)', fontSize: '0.78rem'}}>
                          {Object.values(compareNotes).filter(n => n.trim()).length} note{Object.values(compareNotes).filter(n => n.trim()).length !== 1 ? 's' : ''} added
                        </span>
                      </div>
                      <div style={{display: 'flex', gap: '0.75rem'}}>
                        {Object.keys(compareNotes).length > 0 && (
                          <button
                            className="strategy-tool-btn reset-btn"
                            onClick={() => setCompareNotes({})}
                          >
                            <RotateCcw size={13} /> Clear Notes
                          </button>
                        )}
                        <button
                          className="save-btn"
                          onClick={() => exportCompareNotes(comparisonViewData.asinRows, exportAllRows)}
                          style={{display: 'flex', alignItems: 'center', gap: '0.4rem'}}
                        >
                          <ExternalLink size={15} /> Export .xlsx
                        </button>
                      </div>{/* end right buttons */}
                    </div>{/* end compare-export-bar */}

                  </div>

                ) : (
                  <p style={{color: 'var(--text-secondary)', textAlign: 'center', marginTop: '3rem'}}>
                    Please select both Period A and Period B from the dropdowns above to compare performance.
                  </p>
                )}
              </>
            )}
          </div>
        )}

        {/* TAB 6: BULK OPERATIONS */}
        {activeTab === 'bulk_ops' && (
          <div className="dashboard-container">
            <div className="dashboard-header">
              <div>
                <h1 className="dashboard-title">
                  <ListChecks size={28} color="var(--accent-color)" style={{verticalAlign: 'middle', marginRight: '0.5rem'}} />
                  Bulk Operations Hub
                </h1>
                <p className="date-span-subtitle">
                  Review AI suggestions, assign actions, and export a ready-to-upload Amazon Bulk Operations file.
                </p>
              </div>
              <div style={{display: 'flex', gap: '0.75rem', alignItems: 'center'}}>
                <button className="quick-ai-btn" onClick={aiAutoAssignBulkActions} title="Let AI pre-select actions based on campaign performance">
                  <Bot size={16} color="var(--accent-color)" />
                  <span>AI Auto-Assign</span>
                </button>
                <button className="strategy-tool-btn save-btn" onClick={generateBulkFile} style={{ fontSize: '0.95rem', padding: '0.65rem 1.25rem' }}>
                  ⚡ Export Bulk File
                </button>
              </div>
            </div>

            {(() => {
              const s = getBulkActionSummary();
              return s.total > 0 ? (
                <div className="bulk-summary-bar">
                  <span className="bulk-summary-label">Queued Actions:</span>
                  {s.paused > 0 && <span className="bulk-summary-pill pill-danger">⏸ {s.paused} Paused</span>}
                  {s.enabled > 0 && <span className="bulk-summary-pill pill-success">▶️ {s.enabled} Enabled</span>}
                  {s.scaled > 0 && <span className="bulk-summary-pill pill-info">🚀 {s.scaled} Scaled</span>}
                  {s.reduced > 0 && <span className="bulk-summary-pill pill-warning">📉 {s.reduced} Reduced</span>}
                  <span className="bulk-summary-total">{s.total} total</span>
                </div>
              ) : null;
            })()}

            {rawRecords.length === 0 ? (
              <div className="empty-state">
                <AlertCircle size={48} className="upload-icon" style={{opacity: 0.5}} />
                <p>No campaign data loaded. Please upload a report first.</p>
              </div>
            ) : (
              <div className="campaigns-container">
                <table className="campaign-table bulk-table">
                  <thead>
                    <tr>
                      <th>Campaign Name</th>
                      <th className="text-right">Spend</th>
                      <th className="text-right">Sales</th>
                      <th className="text-right">ACoS</th>
                      <th className="text-right">Budget</th>
                      <th>AI Suggestion</th>
                      <th>Action</th>
                      <th>Notes</th>
                    </tr>
                  </thead>
                  <tbody>
                    {(() => {
                      const campaignsMap = {};
                      rawRecords.forEach(r => {
                        if (!campaignsMap[r.campaignName]) {
                          campaignsMap[r.campaignName] = { name: r.campaignName, spend: 0, sales: 0, budget: r.budget, asin: r.asin };
                        }
                        campaignsMap[r.campaignName].spend += r.spend;
                        campaignsMap[r.campaignName].sales += r.sales;
                        if (r.budget) campaignsMap[r.campaignName].budget = r.budget;
                      });
                      return Object.values(campaignsMap).sort((a,b) => b.spend - a.spend).map((camp, idx) => {
                        const acos = camp.sales > 0 ? (camp.spend / camp.sales) * 100 : (camp.spend > 0 ? 999 : 0);
                        const isBleeding = camp.spend > 0 && camp.sales === 0;
                        
                        let suggestion = "✅ Monitor";
                        let suggestionClass = "suggest-stable";
                        if (camp.spend > globalStrategy.bleedSpendThreshold && camp.sales === 0) {
                          suggestion = "🚨 Bleeding: Pause";
                          suggestionClass = "suggest-danger";
                        } else if (acos > globalStrategy.targetAcos * 1.2) {
                          suggestion = "📉 High ACoS: Reduce";
                          suggestionClass = "suggest-warning";
                        } else if (acos > 0 && acos < globalStrategy.targetAcos * 0.8 && camp.sales > 0) {
                          suggestion = "🚀 Scale Up";
                          suggestionClass = "suggest-success";
                        }

                        const currentAction = bulkActions[camp.name]?.action || 'none';

                        return (
                          <tr key={idx} className={currentAction !== 'none' ? 'bulk-action-active' : ''}>
                            <td className="campaign-name-cell" title={camp.name}>{camp.name}</td>
                            <td className="text-right">{formatCurrency(camp.spend)}</td>
                            <td className="text-right">{formatCurrency(camp.sales)}</td>
                            <td className={`text-right ${isBleeding ? 'bad-acos' : acos > globalStrategy.targetAcos ? 'bad-acos' : 'good-acos'}`}>
                              {isBleeding ? 'N/A ⚠️' : formatPercent(acos)}
                            </td>
                            <td className="text-right">{formatCurrency(camp.budget)}</td>
                            <td><span className={`bulk-suggest ${suggestionClass}`}>{suggestion}</span></td>
                            <td>
                              <select 
                                className="bulk-action-select" 
                                value={currentAction}
                                onChange={(e) => handleBulkActionChange(camp.name, e.target.value, camp.budget)}
                              >
                                <option value="none">-- No Action --</option>
                                <option value="pause">⏸ Pause</option>
                                <option value="enable">▶️ Enable</option>
                                <option value="scale_20">🚀 Scale +20%</option>
                                <option value="scale_30">📈 Scale +30%</option>
                                <option value="scale_50">⚡ Scale +50%</option>
                                <option value="reduce_20">📉 Reduce -20%</option>
                              </select>
                            </td>
                            <td>
                              <input 
                                type="text"
                                className="bulk-note-input"
                                placeholder="Add note..."
                                value={bulkNotes[camp.name] || ''}
                                onChange={(e) => handleBulkNoteChange(camp.name, e.target.value)}
                              />
                            </td>
                          </tr>
                        );
                      });
                    })()}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        )}
        {/* TAB 8: PESH OPTIMIZER */}
        {activeTab === 'pesh_optimizer' && (
          <div className="dashboard-container">
            <div className="dashboard-header">
              <div>
                <h1 className="dashboard-title">
                  <Target size={28} color="var(--purple)" style={{verticalAlign:'middle',marginRight:'0.5rem'}} />
                  PESH Campaign Optimizer
                </h1>
                <p className="date-span-subtitle">
                  AI-driven keyword matrix and waste reduction for PESH strategies.
                </p>
              </div>
            </div>

            {!peshData ? (
              <div className="empty-state">
                <Target size={48} style={{color: 'var(--border-color)', marginBottom: '1rem'}} />
                <h2>No PESH Data Found</h2>
                <p>Ensure your uploaded file contains campaigns with "PESH" in the name, and includes Keyword-level rows (e.g., from a Bulk Operations file).</p>
              </div>
            ) : (
              <>
                {/* EXECUTIVE OVERVIEW */}
                <div className="metrics-grid">
                  <div className="metric-card">
                    <div className="metric-header-row">
                      <span className="metric-title">PESH Spend</span>
                      <TrendingDown size={16} color="var(--text-secondary)" />
                    </div>
                    <div className="metric-value">{formatCurrency(peshData.totalSpend)}</div>
                  </div>
                  <div className="metric-card">
                    <div className="metric-header-row">
                      <span className="metric-title">PESH Sales</span>
                      <TrendingUp size={16} color="var(--text-secondary)" />
                    </div>
                    <div className="metric-value">{formatCurrency(peshData.totalSales)}</div>
                  </div>
                  <div className="metric-card" style={{borderColor: peshData.totalAcos > (globalStrategy.targetAcos/100) ? 'var(--danger)' : 'var(--border-color)'}}>
                    <div className="metric-header-row">
                      <span className="metric-title">PESH ACoS</span>
                      <Target size={16} color="var(--text-secondary)" />
                    </div>
                    <div className="metric-value" style={{color: peshData.totalAcos > (globalStrategy.targetAcos/100) ? 'var(--danger)' : 'inherit'}}>
                      {formatPercent(peshData.totalAcos * 100)}
                    </div>
                  </div>
                  <div className="metric-card">
                    <div className="metric-header-row">
                      <span className="metric-title">Analyzed Targets</span>
                      <BarChart2 size={16} color="var(--text-secondary)" />
                    </div>
                    <div className="metric-value">{peshData.counts.total}</div>
                  </div>
                </div>

                {/* QUADRANT SUMMARY */}
                <div style={{display:'flex', gap:'0.75rem', marginBottom:'1.5rem', flexWrap:'wrap'}}>
                  <span className="strategy-pill" style={{background:'rgba(16, 185, 129, 0.1)', color:'#10b981', border:'1px solid rgba(16, 185, 129, 0.2)'}}>
                    🚀 Top Performers: {peshData.counts.topPerformers}
                  </span>
                  <span className="strategy-pill" style={{background:'rgba(244, 63, 94, 0.1)', color:'#f43f5e', border:'1px solid rgba(244, 63, 94, 0.2)'}}>
                    🛑 Bleeders: {peshData.counts.bleeders}
                  </span>
                  <span className="strategy-pill" style={{background:'rgba(59, 130, 246, 0.1)', color:'#3b82f6', border:'1px solid rgba(59, 130, 246, 0.2)'}}>
                    📈 Under-delivering: {peshData.counts.underDelivering}
                  </span>
                  <span className="strategy-pill" style={{background:'rgba(245, 158, 11, 0.1)', color:'#f59e0b', border:'1px solid rgba(245, 158, 11, 0.2)'}}>
                    ⚠️ Mediocre: {peshData.counts.mediocre}
                  </span>
                </div>

                {/* HIGH PRIORITY WASTE */}
                {peshData.bleeders.length > 0 && (
                  <div style={{background: 'rgba(244, 63, 94, 0.05)', border: '1px solid rgba(244, 63, 94, 0.3)', borderRadius: 'var(--radius-md)', padding: '1.25rem', marginBottom: '2rem'}}>
                    <h3 style={{color: 'var(--danger)', fontSize: '1rem', marginBottom: '1rem', display: 'flex', alignItems: 'center', gap: '0.5rem'}}>
                      <AlertTriangle size={18} /> High-Priority Waste Reductions (Immediate Cuts)
                    </h3>
                    <div style={{maxHeight: '200px', overflowY: 'auto'}}>
                      <table className="campaign-table" style={{background: 'transparent'}}>
                        <thead>
                          <tr>
                            <th style={{background: 'transparent'}}>Keyword</th>
                            <th style={{background: 'transparent'}}>Match</th>
                            <th style={{background: 'transparent'}}>Spend</th>
                            <th style={{background: 'transparent'}}>Sales</th>
                            <th style={{background: 'transparent'}}>Action</th>
                          </tr>
                        </thead>
                        <tbody>
                          {peshData.bleeders.map((b, i) => (
                            <tr key={i} style={{background: 'transparent'}}>
                              <td style={{fontWeight:500}}>{b.keyword}</td>
                              <td>{b.matchType}</td>
                              <td style={{color: 'var(--danger)'}}>{formatCurrency(b.spend)}</td>
                              <td>{formatCurrency(b.sales)}</td>
                              <td><span className="strategy-pill" style={{background:'rgba(244, 63, 94, 0.1)', color:'#f43f5e'}}>{b.action}</span></td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}

                {/* KEYWORD OPTIMIZATION MATRIX */}
                <h3 style={{fontSize: '1.1rem', marginBottom: '1rem', color: 'var(--text-primary)'}}>Keyword Optimization Matrix</h3>
                <div className="comparison-table-container">
                  <table className="campaign-table">
                    <thead>
                      <tr>
                        <th>Campaign</th>
                        <th>Ad Group</th>
                        <th>Keyword</th>
                        <th>Match</th>
                        <th className="text-right">Bid</th>
                        <th className="text-right">Clicks</th>
                        <th className="text-right">Spend</th>
                        <th className="text-right">Sales</th>
                        <th className="text-right">Orders</th>
                        <th className="text-right">ACoS</th>
                        <th className="text-right">Suggested Bid</th>
                        <th>Recommended Action</th>
                        <th>Rationale</th>
                      </tr>
                    </thead>
                    <tbody>
                      {peshData.records.map((row, i) => (
                        <tr key={i}>
                          <td style={{maxWidth: '150px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap'}} title={row.campaign}>{row.campaign}</td>
                          <td style={{maxWidth: '120px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap'}} title={row.adGroup}>{row.adGroup}</td>
                          <td style={{fontWeight: 600}}>{row.keyword}</td>
                          <td>{row.matchType}</td>
                          <td className="text-right">{row.bid !== null ? formatCurrency(row.bid) : '-'}</td>
                          <td className="text-right">{row.clicks}</td>
                          <td className="text-right">{formatCurrency(row.spend)}</td>
                          <td className="text-right">{formatCurrency(row.sales)}</td>
                          <td className="text-right">{row.orders}</td>
                          <td className="text-right" style={{color: row.acos > (globalStrategy.targetAcos/100) ? 'var(--danger)' : 'inherit'}}>
                            {formatPercent(row.acos * 100)}
                          </td>
                          <td className="text-right" style={{fontWeight: 'bold', color: 'var(--accent-color)'}}>
                            {formatCurrency(row.suggestedBid)}
                          </td>
                          <td>
                            <span className="strategy-pill" style={{
                              background: row.quadrant === 'Bleeder' ? 'rgba(244, 63, 94, 0.1)' : row.quadrant === 'Top Performer' ? 'rgba(16, 185, 129, 0.1)' : row.quadrant === 'Under-delivering' ? 'rgba(59, 130, 246, 0.1)' : 'var(--bg-tertiary)',
                              color: row.quadrant === 'Bleeder' ? '#f43f5e' : row.quadrant === 'Top Performer' ? '#10b981' : row.quadrant === 'Under-delivering' ? '#3b82f6' : 'var(--text-primary)',
                            }}>
                              {row.action}
                            </span>
                          </td>
                          <td style={{fontSize: '0.8rem', color: 'var(--text-secondary)'}}>{row.rationale}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </>
            )}
          </div>
        )}


        {/* TAB 7: ADMIN PANEL */}
        {activeTab === 'admin' && (() => {
          const typeLabels = { ui: '🎨 UI/UX', bug: '🐛 Bug', perf: '⚡ Performance', feature: '💡 Feature', other: '📝 Other' };
          const openItems = feedbacks.filter(f => f.status === 'open');
          const resolvedItems = feedbacks.filter(f => f.status === 'resolved');
          return (
            <div className="dashboard-container">
              <div className="dashboard-header">
                <div>
                  <h1 className="dashboard-title">
                    <ShieldCheck size={28} color="var(--accent-color)" style={{verticalAlign:'middle',marginRight:'0.5rem'}} />
                    Admin Panel — Feedback Dashboard
                  </h1>
                  <p style={{color:'var(--text-secondary)',fontSize:'0.9rem',marginTop:'0.3rem'}}>
                    All user feedback is stored locally. Clusters detect repeated issues. Auto-implement triggers at 3+ matching reports.
                  </p>
                </div>
                <div style={{display:'flex',gap:'0.75rem',alignItems:'center'}}>
                  <span style={{color:'var(--text-secondary)',fontSize:'0.85rem'}}>
                    {feedbacks.length} total · {openItems.length} open · {resolvedItems.length} resolved
                  </span>
                  {feedbacks.length > 0 && (
                    <button className="strategy-tool-btn reset-btn" onClick={() => { if(window.confirm('Clear ALL feedback? This cannot be undone.')) setFeedbacks([]); }}>
                      <Trash2 size={14} /> Clear All
                    </button>
                  )}
                </div>
              </div>

              {feedbackClusters.filter(c => c.count >= 3).length > 0 && (
                <div style={{marginBottom:'1.5rem'}}>
                  <h3 style={{fontSize:'1rem',marginBottom:'0.75rem',display:'flex',alignItems:'center',gap:'0.5rem'}}>
                    <Zap size={18} color="var(--warning)" /> Auto-Implement Triggers
                    <span style={{fontSize:'0.78rem',color:'var(--text-secondary)',fontWeight:400}}>— 3+ users reported the same issue</span>
                  </h3>
                  {feedbackClusters.filter(c => c.count >= 3).map((cluster, ci) => (
                    <div key={ci} className="admin-auto-implement-banner">
                      <div className="auto-implement-left">
                        <span className="auto-implement-count">{cluster.count}x</span>
                        <div>
                          <div className="auto-implement-title">Common issue: <strong>{cluster.topKeywords.slice(0,3).join(', ')}</strong></div>
                          <div className="auto-implement-keywords">
                            {cluster.topKeywords.map((kw,ki) => <span key={ki} className="keyword-pill">{kw}</span>)}
                          </div>
                          <div style={{marginTop:'0.4rem',fontSize:'0.8rem',color:'var(--text-secondary)'}}>
                            "{cluster.items[0].text.slice(0,120)}{cluster.items[0].text.length > 120 ? '...' : ''}"
                          </div>
                        </div>
                      </div>
                      <div className="auto-implement-actions">
                        <span className="auto-implement-ready-badge">⚡ Auto-Implement Ready</span>
                        <button className="strategy-tool-btn apply-all-btn" onClick={() => cluster.items.forEach(item => toggleFeedbackStatus(item.id))}>
                          <CheckCircle2 size={14} /> Mark All Resolved
                        </button>
                      </div>
                    </div>
                  ))}
                </div>
              )}

              {feedbackClusters.filter(c => c.count === 2).length > 0 && (
                <div style={{marginBottom:'1.5rem'}}>
                  <h3 style={{fontSize:'1rem',marginBottom:'0.75rem',display:'flex',alignItems:'center',gap:'0.5rem'}}>
                    <AlertTriangle size={18} color="var(--warning)" /> Emerging Patterns
                    <span style={{fontSize:'0.78rem',color:'var(--text-secondary)',fontWeight:400}}>— 1 more report needed to trigger auto-implement</span>
                  </h3>
                  {feedbackClusters.filter(c => c.count === 2).map((cluster, ci) => (
                    <div key={ci} className="admin-cluster-card">
                      <div style={{display:'flex',alignItems:'center',gap:'0.75rem',flexWrap:'wrap'}}>
                        <span className="cluster-count-badge">{cluster.count} reports</span>
                        <div className="auto-implement-keywords">
                          {cluster.topKeywords.map((kw,ki) => <span key={ki} className="keyword-pill">{kw}</span>)}
                        </div>
                      </div>
                      <div style={{color:'var(--text-secondary)',fontSize:'0.8rem',marginTop:'0.35rem'}}>
                        "{cluster.items[0].text.slice(0,100)}..."
                      </div>
                    </div>
                  ))}
                </div>
              )}

              <h3 style={{fontSize:'1rem',marginBottom:'0.75rem',display:'flex',alignItems:'center',gap:'0.5rem'}}>
                <MessageSquare size={18} color="var(--accent-color)" /> Feedback Inbox
              </h3>

              {feedbacks.length === 0 ? (
                <div className="empty-state">
                  <MessageSquare size={48} className="empty-icon" />
                  <p>No feedback yet. Share the app and responses will appear here.</p>
                </div>
              ) : (
                <div className="campaigns-container">
                  <table className="campaign-table">
                    <thead>
                      <tr>
                        <th style={{width:'110px'}}>Date</th>
                        <th style={{width:'120px'}}>Type</th>
                        <th>Message</th>
                        <th style={{width:'130px'}}>Keywords</th>
                        <th style={{width:'100px'}}>Status</th>
                        <th style={{width:'80px'}}>Actions</th>
                      </tr>
                    </thead>
                    <tbody>
                      {feedbacks.map((fb) => (
                        <tr key={fb.id} className={fb.status === 'resolved' ? 'admin-row-resolved' : ''}>
                          <td style={{fontSize:'0.78rem',color:'var(--text-secondary)'}}>
                            {new Date(fb.timestamp).toLocaleDateString()}<br/>
                            <span style={{fontSize:'0.72rem'}}>{new Date(fb.timestamp).toLocaleTimeString([],{hour:'2-digit',minute:'2-digit'})}</span>
                          </td>
                          <td>
                            <span className="perf-badge" style={
                              fb.type === 'bug' ? {background:'rgba(239,68,68,0.15)',color:'#f87171',border:'1px solid rgba(239,68,68,0.3)'} :
                              fb.type === 'feature' ? {background:'rgba(16,185,129,0.15)',color:'#34d399',border:'1px solid rgba(16,185,129,0.3)'} :
                              fb.type === 'perf' ? {background:'rgba(245,158,11,0.15)',color:'#fbbf24',border:'1px solid rgba(245,158,11,0.3)'} :
                              {background:'rgba(59,130,246,0.15)',color:'#60a5fa',border:'1px solid rgba(59,130,246,0.3)'}
                            }>
                              {typeLabels[fb.type] || fb.type}
                            </span>
                          </td>
                          <td style={{fontSize:'0.85rem'}}>
                            <span className={fb.status === 'resolved' ? 'admin-text-resolved' : ''}>{fb.text}</span>
                          </td>
                          <td>
                            <div style={{display:'flex',flexWrap:'wrap',gap:'0.2rem'}}>
                              {fb.keywords.slice(0,4).map((kw,ki) => <span key={ki} className="keyword-pill keyword-pill-sm">{kw}</span>)}
                            </div>
                          </td>
                          <td>
                            <span className="perf-badge" style={fb.status === 'resolved'
                              ? {background:'rgba(16,185,129,0.15)',color:'#34d399',border:'1px solid rgba(16,185,129,0.3)'}
                              : {background:'rgba(239,68,68,0.15)',color:'#f87171',border:'1px solid rgba(239,68,68,0.3)'}}>
                              {fb.status === 'resolved' ? '✓ Resolved' : '● Open'}
                            </span>
                          </td>
                          <td>
                            <div style={{display:'flex',gap:'0.4rem',alignItems:'center'}}>
                              <button className="compare-ai-action-btn" title={fb.status==='resolved'?'Reopen':'Mark resolved'}
                                onClick={() => toggleFeedbackStatus(fb.id)} style={{padding:'0.25rem 0.5rem',fontSize:'0.75rem'}}>
                                {fb.status === 'resolved' ? <RotateCcw size={13}/> : <CheckCircle2 size={13}/>}
                              </button>
                              <button className="compare-ai-action-btn" title="Delete"
                                onClick={() => deleteFeedback(fb.id)} style={{padding:'0.25rem 0.5rem',fontSize:'0.75rem',borderColor:'var(--danger)',color:'var(--danger)'}}>
                                <Trash2 size={13}/>
                              </button>
                            </div>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          );
        })()}

        {/* MODAL: ASIN AI STRATEGY & ACTION AUDIT */}
        {selectedAsinForAI && (
          <div className="modal-overlay" onClick={() => setSelectedAsinForAI(null)}>
            <div className="modal-content ai-modal" onClick={(e) => e.stopPropagation()}>
              <div className="modal-header">
                <div style={{display: 'flex', alignItems: 'center', gap: '0.6rem'}}>
                  <Bot size={24} color="var(--accent-color)" />
                  <div>
                    <h2>AI PPC Growth & Strategy Action Plan</h2>
                    <span className="modal-sub">
                      Autonomous Decision Engine for <strong>{selectedAsinForAI.plan.productIdentity.asin}</strong>
                    </span>
                  </div>
                </div>
                <button className="close-modal-btn" onClick={() => setSelectedAsinForAI(null)}>
                  <X size={20} />
                </button>
              </div>

              <div className="modal-body">
                {/* 1. Product Identity */}
                <div className="ai-section-box">
                  <div className="ai-section-title">
                    <Search size={16} color="var(--accent-color)" />
                    <h4>1. Product Identity & Context</h4>
                  </div>
                  <div className="ai-identity-grid">
                    <div><strong>Product:</strong> {selectedAsinForAI.plan.productIdentity.productName}</div>
                    <div><strong>ASIN:</strong> {selectedAsinForAI.plan.productIdentity.asin}</div>
                    <div><strong>Brand:</strong> {selectedAsinForAI.plan.productIdentity.brand}</div>
                    <div><strong>Category:</strong> {selectedAsinForAI.plan.productIdentity.category}</div>
                  </div>
                </div>

                {/* 2. Strategy Applied */}
                <div className="ai-section-box">
                  <div className="ai-section-title">
                    <Target size={16} color="var(--accent-color)" />
                    <h4>2. Applied Strategy & Alignment</h4>
                  </div>
                  <div className="ai-strategy-applied-badge">
                    <span className={`strategy-pill ${selectedAsinForAI.plan.strategyConfig.isCustomOverride ? 'custom-pill' : 'global-pill'}`}>
                      {selectedAsinForAI.plan.strategyConfig.isCustomOverride ? 'ASIN-Specific Override' : 'Global Portfolio Strategy'}
                    </span>
                    <span style={{fontWeight: 600}}>{selectedAsinForAI.plan.strategyConfig.modeLabel}</span>
                  </div>
                  <div className="ai-kpi-summary-strip">
                    <div>Target ACoS: <strong>{selectedAsinForAI.plan.strategyConfig.targetAcos}%</strong></div>
                    <div>Actual ACoS: <strong className={selectedAsinForAI.plan.performanceMetrics.acos > selectedAsinForAI.plan.strategyConfig.targetAcos ? 'bad-acos' : 'good-acos'}>{formatPercent(selectedAsinForAI.plan.performanceMetrics.acos)}</strong></div>
                    <div>Spend: <strong>{formatCurrency(selectedAsinForAI.plan.performanceMetrics.spend)}</strong></div>
                    <div>Sales: <strong>{formatCurrency(selectedAsinForAI.plan.performanceMetrics.sales)}</strong></div>
                    <div>CVR: <strong>{formatPercent(selectedAsinForAI.plan.performanceMetrics.cvr)}</strong></div>
                  </div>
                </div>

                {/* 3. Actionable Optimizations */}
                <div className="ai-section-box">
                  <div className="ai-section-title">
                    <Zap size={16} color="var(--accent-color)" />
                    <h4>3. Step-by-Step Optimization Actions</h4>
                  </div>
                  <div className="ai-actions-list">
                    {selectedAsinForAI.plan.actions.map((act, idx) => (
                      <div key={idx} className={`ai-action-card action-level-${act.level}`}>
                        <div className="action-card-header">
                          <span className="action-number">Action #{idx + 1}</span>
                          <span className="action-target">{act.target}</span>
                        </div>
                        <p className="action-directive">{act.action}</p>
                        <span className="action-impact">💡 {act.impact}</span>
                      </div>
                    ))}
                  </div>
                </div>

                {/* 4. Strategic Rationale */}
                <div className="ai-section-box">
                  <div className="ai-section-title">
                    <Cpu size={16} color="var(--accent-color)" />
                    <h4>4. Data-Driven Strategic Rationale</h4>
                  </div>
                  <ul className="ai-rationale-list">
                    {selectedAsinForAI.plan.rationales.map((rat, i) => (
                      <li key={i}>{rat}</li>
                    ))}
                  </ul>
                </div>
              </div>

              <div className="modal-footer">
                <button className="copy-plan-btn" onClick={() => copyAIPlanText(selectedAsinForAI.plan)}>
                  {copiedActionMsg ? <Check size={16} color="#10b981" /> : <Copy size={16} />}
                  {copiedActionMsg ? 'Plan Copied to Clipboard!' : 'Copy Full Action Plan'}
                </button>
                <button className="close-action-btn" onClick={() => setSelectedAsinForAI(null)}>
                  Close
                </button>
              </div>
            </div>
          </div>
        )}

        {/* MODAL: COMPARATIVE AI STRATEGY & WEEK-OVER-WEEK ACTION AUDIT */}
        {selectedCompareAsinForAI && (
          <div className="modal-overlay" onClick={() => setSelectedCompareAsinForAI(null)}>
            <div className="modal-content ai-modal" onClick={(e) => e.stopPropagation()}>
              <div className="modal-header">
                <div style={{display: 'flex', alignItems: 'center', gap: '0.6rem'}}>
                  <Bot size={24} color="var(--accent-color)" />
                  <div>
                    <h2>AI Week-over-Week Strategy Directive</h2>
                    <span className="modal-sub">
                      Comparative Growth Analysis for <strong>{selectedCompareAsinForAI.report.title}</strong>
                    </span>
                  </div>
                </div>
                <button className="close-modal-btn" onClick={() => setSelectedCompareAsinForAI(null)}>
                  <X size={20} />
                </button>
              </div>

              <div className="modal-body">
                {/* 1. Product Identity */}
                <div className="ai-section-box">
                  <div className="ai-section-title">
                    <Search size={16} color="var(--accent-color)" />
                    <h4>1. Product Identity & Strategic Context</h4>
                  </div>
                  <div className="ai-identity-grid">
                    <div><strong>Product:</strong> {selectedCompareAsinForAI.report.title}</div>
                    <div><strong>ASIN:</strong> {selectedCompareAsinForAI.report.asin}</div>
                    <div><strong>Brand:</strong> {selectedCompareAsinForAI.report.brand}</div>
                    <div><strong>Category:</strong> {selectedCompareAsinForAI.report.category}</div>
                    <div><strong>Applied Mode:</strong> <span className={`strategy-pill mode-${selectedCompareAsinForAI.report.mode}`}>{selectedCompareAsinForAI.report.mode.toUpperCase()}</span></div>
                    <div><strong>Target ACoS:</strong> <strong>{selectedCompareAsinForAI.report.targetAcos}%</strong></div>
                  </div>
                </div>

                {/* 2. Side-by-Side Period Delta */}
                <div className="ai-section-box">
                  <div className="ai-section-title">
                    <GitCompare size={16} color="var(--accent-color)" />
                    <h4>2. Comparative Performance Evolution</h4>
                  </div>
                  <div className="ai-kpi-summary-strip" style={{display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(130px, 1fr))', gap: '0.75rem'}}>
                    <div>
                      <span style={{fontSize: '0.7rem', color: 'var(--text-secondary)', display: 'block'}}>SPEND DELTA</span>
                      <strong>{formatCurrency(selectedCompareAsinForAI.report.spendB)}</strong>
                      <span style={{fontSize: '0.75rem', color: 'var(--text-secondary)', display: 'block'}}>
                        was {formatCurrency(selectedCompareAsinForAI.report.spendA)} ({selectedCompareAsinForAI.report.spendDiff >= 0 ? '+' : ''}{selectedCompareAsinForAI.report.spendPct.toFixed(1)}%)
                      </span>
                    </div>

                    <div>
                      <span style={{fontSize: '0.7rem', color: 'var(--text-secondary)', display: 'block'}}>SALES DELTA</span>
                      <strong>{formatCurrency(selectedCompareAsinForAI.report.salesB)}</strong>
                      <span style={{fontSize: '0.75rem', color: 'var(--text-secondary)', display: 'block'}}>
                        was {formatCurrency(selectedCompareAsinForAI.report.salesA)} ({selectedCompareAsinForAI.report.salesDiff >= 0 ? '+' : ''}{selectedCompareAsinForAI.report.salesPct.toFixed(1)}%)
                      </span>
                    </div>

                    <div>
                      <span style={{fontSize: '0.7rem', color: 'var(--text-secondary)', display: 'block'}}>ACOS DELTA</span>
                      <strong className={selectedCompareAsinForAI.report.acosB > selectedCompareAsinForAI.report.targetAcos ? 'bad-acos' : 'good-acos'}>
                        {formatPercent(selectedCompareAsinForAI.report.acosB)}
                      </strong>
                      <span style={{fontSize: '0.75rem', color: 'var(--text-secondary)', display: 'block'}}>
                        was {formatPercent(selectedCompareAsinForAI.report.acosA)} ({selectedCompareAsinForAI.report.acosDiff >= 0 ? '+' : ''}{selectedCompareAsinForAI.report.acosDiff.toFixed(2)}%)
                      </span>
                    </div>

                    <div>
                      <span style={{fontSize: '0.7rem', color: 'var(--text-secondary)', display: 'block'}}>ORDERS DELTA</span>
                      <strong>{selectedCompareAsinForAI.report.ordersB} orders</strong>
                      <span style={{fontSize: '0.75rem', color: 'var(--text-secondary)', display: 'block'}}>
                        was {selectedCompareAsinForAI.report.ordersA} ({selectedCompareAsinForAI.report.ordersDiff >= 0 ? '+' : ''}{selectedCompareAsinForAI.report.ordersDiff})
                      </span>
                    </div>
                  </div>
                </div>

                {/* 3. AI Strategic Diagnosis */}
                <div className="ai-section-box">
                  <div className="ai-section-title">
                    <Cpu size={16} color="var(--accent-color)" />
                    <h4>3. AI Week-over-Week Diagnosis</h4>
                  </div>
                  <div style={{display: 'flex', alignItems: 'center', gap: '0.6rem', marginBottom: '0.5rem'}}>
                    <span className={`compare-ai-badge badge-${selectedCompareAsinForAI.report.badge.type}`}>
                      {selectedCompareAsinForAI.report.badge.label}
                    </span>
                  </div>
                  <p style={{fontSize: '0.9rem', color: 'var(--text-primary)', lineHeight: 1.5, margin: 0}}>
                    {selectedCompareAsinForAI.report.rationale}
                  </p>
                </div>

                {/* 4. Action Directives */}
                <div className="ai-section-box">
                  <div className="ai-section-title">
                    <Zap size={16} color="var(--accent-color)" />
                    <h4>4. Recommended Campaign Manager Actions (This Week)</h4>
                  </div>
                  <div className="ai-actions-list">
                    {selectedCompareAsinForAI.report.actionDirectives.map((act, idx) => (
                      <div key={idx} className={`ai-action-card action-level-${selectedCompareAsinForAI.report.priority === 1 ? 'danger' : selectedCompareAsinForAI.report.priority === 2 ? 'success' : 'info'}`}>
                        <div className="action-card-header">
                          <span className="action-number">Directive #{idx + 1}</span>
                          <span className="action-target">{selectedCompareAsinForAI.report.asin}</span>
                        </div>
                        <p className="action-directive">{act}</p>
                      </div>
                    ))}
                  </div>
                </div>
              </div>

              <div className="modal-footer">
                <button className="copy-plan-btn" onClick={() => copyCompareAIPlanText(selectedCompareAsinForAI)}>
                  {copiedActionMsg ? <Check size={16} color="#10b981" /> : <Copy size={16} />}
                  {copiedActionMsg ? 'Directives Copied to Clipboard!' : 'Copy Directive Instructions'}
                </button>
                <button className="close-action-btn" onClick={() => setSelectedCompareAsinForAI(null)}>
                  Close
                </button>
              </div>
            </div>
          </div>
        )}

        {/* MODAL: ASIN CATALOG & CUSTOM STRATEGY EDITOR */}
        {editingCatalogAsin && (
          <div className="modal-overlay" onClick={() => setEditingCatalogAsin(null)}>
            <div className="modal-content catalog-modal" onClick={(e) => e.stopPropagation()}>
              <div className="modal-header">
                <div>
                  <h2>ASIN Catalog & Custom Strategy</h2>
                  <span className="modal-sub">Configure attributes and custom override directives for <strong>{editingCatalogAsin}</strong></span>
                </div>
                <button className="close-modal-btn" onClick={() => setEditingCatalogAsin(null)}>
                  <X size={20} />
                </button>
              </div>

              <div className="modal-body">
                <div className="form-group">
                  <div style={{display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '0.25rem'}}>
                    <label style={{margin: 0}}>Product Descriptive Title / Name</label>
                    <button 
                      type="button" 
                      className="ai-detect-title-btn"
                      onClick={autoDetectCurrentAsinTitle}
                      title="Analyze campaign names using AI to extract the product title automatically"
                    >
                      <Sparkles size={13} />
                      <span>AI Auto-Detect Name</span>
                    </button>
                  </div>
                  <input 
                    type="text" 
                    value={editForm.productName} 
                    onChange={(e) => setEditForm({ ...editForm, productName: e.target.value })}
                    className="form-control"
                    placeholder="e.g. Stainless Steel French Press Coffee Maker 34oz"
                  />
                </div>

                <div className="settings-form-grid" style={{marginBottom: '1rem'}}>
                  <div className="form-group">
                    <label>Brand Name</label>
                    <input 
                      type="text" 
                      value={editForm.brand} 
                      onChange={(e) => setEditForm({ ...editForm, brand: e.target.value })}
                      className="form-control"
                      placeholder="e.g. CoffeeCraft"
                    />
                  </div>

                  <div className="form-group">
                    <label>Category</label>
                    <input 
                      type="text" 
                      value={editForm.category} 
                      onChange={(e) => setEditForm({ ...editForm, category: e.target.value })}
                      className="form-control"
                      placeholder="e.g. Home & Kitchen"
                    />
                  </div>
                </div>

                <div className="settings-form-grid" style={{marginBottom: '1rem'}}>
                  <div className="form-group">
                    <label>Strategy Mode Override</label>
                    <select 
                      value={editForm.strategyMode} 
                      onChange={(e) => setEditForm({ ...editForm, strategyMode: e.target.value })}
                      className="form-control"
                    >
                      <option value="global">Inherit Global Portfolio Strategy</option>
                      <option value="launch">Launch / Cold Start (Aggressive Discovery)</option>
                      <option value="growth">Growth / Scale (Harvesting & Ranking)</option>
                      <option value="profitability">Profitability / Margin Protection</option>
                      <option value="defense">Defense (Protect Brand & PDPs)</option>
                      <option value="custom">Strict Custom Directives</option>
                    </select>
                  </div>

                  <div className="form-group">
                    <label>Target ACoS for this ASIN (%)</label>
                    <div className="stepper-input-wrapper">
                      <button 
                        type="button" 
                        className="stepper-btn"
                        onClick={() => setEditForm(prev => ({ ...prev, targetAcos: Math.max(1, (parseFloat(prev.targetAcos) || 30) - 1) }))}
                        title="Decrease by 1%"
                      >
                        <Minus size={14} />
                      </button>
                      <input 
                        type="number" 
                        step="1"
                        min="1"
                        max="100"
                        value={editForm.targetAcos ?? ''} 
                        onChange={(e) => setEditForm({ ...editForm, targetAcos: e.target.value })}
                        onBlur={() => setEditForm(prev => ({ ...prev, targetAcos: parseFloat(prev.targetAcos) || globalStrategy.targetAcos }))}
                        className="form-control stepper-input"
                      />
                      <button 
                        type="button" 
                        className="stepper-btn"
                        onClick={() => setEditForm(prev => ({ ...prev, targetAcos: Math.min(100, (parseFloat(prev.targetAcos) || 30) + 1) }))}
                        title="Increase by 1%"
                      >
                        <Plus size={14} />
                      </button>
                    </div>
                  </div>
                </div>

                <div className="form-group">
                  <label>Custom Strategy Instructions / Overrides</label>
                  <textarea 
                    value={editForm.customDirectives} 
                    onChange={(e) => setEditForm({ ...editForm, customDirectives: e.target.value })}
                    className="form-control"
                    rows="3"
                    placeholder="e.g. Tolerate up to 50% ACoS to push ranking on 'espresso machine'. Do not negate competitor brand terms."
                  ></textarea>
                </div>
              </div>

              <div className="modal-footer">
                <button className="save-snapshot-btn" onClick={saveCatalogEntry}>
                  <Save size={16} /> Save ASIN Configuration
                </button>
                <button className="close-action-btn" onClick={() => setEditingCatalogAsin(null)}>
                  Cancel
                </button>
              </div>
            </div>
          </div>
        )}
        {/* MODAL: FEEDBACK */}
        {isFeedbackOpen && (
          <div className="modal-overlay" onClick={() => { setIsFeedbackOpen(false); setFeedbackSubmitted(false); setFeedbackText(''); }}>
            <div className="modal-content" onClick={(e) => e.stopPropagation()} style={{ maxWidth: '500px' }}>
              <div className="modal-header">
                <h3>Submit Feedback</h3>
                <button className="close-btn" onClick={() => { setIsFeedbackOpen(false); setFeedbackSubmitted(false); setFeedbackText(''); }}>
                  <X size={18} />
                </button>
              </div>
              <div className="modal-body" style={{ padding: '1.5rem' }}>
                {feedbackSubmitted ? (
                  <div style={{ textAlign: 'center', padding: '2rem 0' }}>
                    <CheckCircle2 size={48} color="#10b981" style={{ margin: '0 auto 1rem' }} />
                    <h4>Thank you for your feedback!</h4>
                    <p style={{ color: 'var(--text-secondary)', marginTop: '0.5rem' }}>
                      Your input helps us improve the Campaign Manager Analyzer.
                    </p>
                    <button 
                      className="save-btn" 
                      style={{ marginTop: '1.5rem' }}
                      onClick={() => { setIsFeedbackOpen(false); setFeedbackSubmitted(false); setFeedbackText(''); }}
                    >
                      Close
                    </button>
                  </div>
                ) : (
                  <>
                    <div className="form-group" style={{ marginBottom: '1.2rem' }}>
                      <label>Feedback Category</label>
                      <select 
                        className="form-control"
                        value={feedbackType}
                        onChange={(e) => setFeedbackType(e.target.value)}
                      >
                        <option value="ui">🎨 UI / UX Enhancement</option>
                        <option value="bug">🐛 Bug Report</option>
                        <option value="perf">⚡ Performance Issue</option>
                        <option value="feature">💡 Feature Request</option>
                        <option value="other">📝 Other</option>
                      </select>
                    </div>
                    <div className="form-group">
                      <label>Description</label>
                      <textarea 
                        className="form-control" 
                        rows={5}
                        placeholder="Please describe your feedback in detail..."
                        value={feedbackText}
                        onChange={(e) => setFeedbackText(e.target.value)}
                        style={{ resize: 'vertical' }}
                      />
                    </div>
                    <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '1rem', marginTop: '1.5rem' }}>
                      <button 
                        className="cancel-btn"
                        onClick={() => setIsFeedbackOpen(false)}
                      >
                        Cancel
                      </button>
                      <button 
                        className="save-btn"
                        onClick={submitFeedback}
                        disabled={!feedbackText.trim()}
                      >
                        <Send size={16} /> Submit Feedback
                      </button>
                    </div>
                  </>
                )}
              </div>
            </div>
          </div>
        )}
      </main>
    </div>
  );
}

export default App;
