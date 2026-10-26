#!/usr/bin/env node
/**
 * validate-coverage-ledger.cjs — Zero-dependency validator for coverage-ledger.json
 * Enforces Cloudflare Security Audit coverage tracking contract.
 */
const fs = require('fs');
const path = require('path');

function validateCoverageLedger(filePath) {
  if (!fs.existsSync(filePath)) {
    console.error(`[ERROR] Coverage ledger not found: ${filePath}`);
    process.exit(1);
  }

  let data;
  try {
    data = JSON.parse(fs.readFileSync(filePath, 'utf8'));
  } catch (err) {
    console.error(`[ERROR] Invalid JSON in ${filePath}: ${err.message}`);
    process.exit(1);
  }

  if (!data.version || !data.target || !Array.isArray(data.units)) {
    console.error(`[ERROR] Missing required fields: 'version', 'target', or 'units' array.`);
    process.exit(1);
  }

  const validStatuses = ['unreviewed', 'in_progress', 'covered'];
  for (const unit of data.units) {
    if (!unit.id || !unit.file || !unit.status) {
      console.error(`[ERROR] Unit ${unit.id || 'unknown'} missing required fields ('id', 'file', 'status')`);
      process.exit(1);
    }
    if (!validStatuses.includes(unit.status)) {
      console.error(`[ERROR] Invalid status '${unit.status}' in unit ${unit.id}`);
      process.exit(1);
    }
  }

  const coveredCount = data.units.filter(u => u.status === 'covered').length;
  console.log(`[OK] coverage-ledger.json is valid: ${coveredCount}/${data.units.length} units covered (${Math.round(coveredCount / data.units.length * 100)}%).`);
  process.exit(0);
}

const targetFile = process.argv[2] || path.join(process.cwd(), 'coverage-ledger.json');
validateCoverageLedger(targetFile);
