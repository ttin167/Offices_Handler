#!/usr/bin/env node
/**
 * validate-findings.cjs — Zero-dependency validator for findings.json
 * Enforces Cloudflare Security Audit findings contract.
 */
const fs = require('fs');
const path = require('path');

function validateFindings(filePath) {
  if (!fs.existsSync(filePath)) {
    console.error(`[ERROR] Findings file not found: ${filePath}`);
    process.exit(1);
  }

  let data;
  try {
    data = JSON.parse(fs.readFileSync(filePath, 'utf8'));
  } catch (err) {
    console.error(`[ERROR] Invalid JSON in ${filePath}: ${err.message}`);
    process.exit(1);
  }

  const requiredTop = ['version', 'target', 'timestamp', 'findings'];
  for (const field of requiredTop) {
    if (!data[field]) {
      console.error(`[ERROR] Missing required top-level field: '${field}'`);
      process.exit(1);
    }
  }

  const findings = data.findings;
  const categories = ['confirmed', 'needs_validation', 'rejected'];
  for (const cat of categories) {
    if (!Array.isArray(findings[cat])) {
      console.error(`[ERROR] findings.${cat} must be an array`);
      process.exit(1);
    }
  }

  // Validate confirmed
  for (const item of findings.confirmed) {
    const required = ['id', 'title', 'severity', 'attack_class', 'file', 'line_range', 'evidence', 'fix_proposal'];
    for (const r of required) {
      if (!item[r]) {
        console.error(`[ERROR] Confirmed item ${item.id || 'unknown'} missing required field '${r}'`);
        process.exit(1);
      }
    }
    const validSeverities = ['CRITICAL', 'HIGH', 'MEDIUM', 'LOW', 'INFO'];
    if (!validSeverities.includes(item.severity)) {
      console.error(`[ERROR] Invalid severity '${item.severity}' in item ${item.id}`);
      process.exit(1);
    }
  }

  // Validate needs_validation
  for (const item of findings.needs_validation) {
    const required = ['id', 'title', 'attack_class', 'file', 'unresolved_fact', 'validation_plan'];
    for (const r of required) {
      if (!item[r]) {
        console.error(`[ERROR] Needs-validation item ${item.id || 'unknown'} missing required field '${r}'`);
        process.exit(1);
      }
    }
    if (item.severity) {
      console.error(`[ERROR] needs_validation item ${item.id} must NOT declare a severity score`);
      process.exit(1);
    }
  }

  // Validate rejected
  for (const item of findings.rejected) {
    const required = ['id', 'title', 'attack_class', 'file', 'disprove_rationale'];
    for (const r of required) {
      if (!item[r]) {
        console.error(`[ERROR] Rejected item ${item.id || 'unknown'} missing required field '${r}'`);
        process.exit(1);
      }
    }
  }

  console.log(`[OK] findings.json is structurally valid: ${findings.confirmed.length} confirmed, ${findings.needs_validation.length} needs_validation, ${findings.rejected.length} rejected.`);
  process.exit(0);
}

const targetFile = process.argv[2] || path.join(process.cwd(), 'findings.json');
validateFindings(targetFile);
