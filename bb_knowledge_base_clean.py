"""
🔥 BLAZE BUDDY v3.0
Cannabis Inventory Intelligence System
Clean, Modern, Production-Ready Implementation
"""

import sqlite3
import pandas as pd
import hashlib
import json
import logging
from datetime import datetime
from difflib import SequenceMatcher
from pathlib import Path
from typing import Tuple, Dict, List, Optional

# Configure logging
logging.basicConfig(level=logging.INFO, format='[%(asctime)s] %(levelname)s: %(message)s')
logger = logging.getLogger(__name__)


class KnowledgeBase:
    """
    The intelligent brain that learns from every file.
    
    Features:
    - Auto-detect file types (government, inventory, sales, PO)
    - Link messy SKUs to official government IDs
    - Build time-series analytics
    - Grow smarter with every import
    """
    
    def __init__(self, db_path: str = "blaze_buddy.db"):
        self.db_path = db_path
        self.db = sqlite3.connect(db_path)
        self.cursor = self.db.cursor()
        self._bootstrap()
    
    def _bootstrap(self):
        """Create database schema"""
        schema = """
        CREATE TABLE IF NOT EXISTS master_catalog (
            gov_id TEXT PRIMARY KEY,
            product_name TEXT NOT NULL,
            category_level1 TEXT,
            category_level2 TEXT,
            category_level3 TEXT,
            wholesale_cost REAL,
            case_size INTEGER DEFAULT 1,
            thc_min REAL, thc_max REAL,
            cbd_min REAL, cbd_max REAL,
            last_updated TIMESTAMP,
            data_source TEXT
        );
        
        CREATE TABLE IF NOT EXISTS sku_linkage (
            linkage_id INTEGER PRIMARY KEY,
            greenline_sku TEXT UNIQUE NOT NULL,
            gov_id TEXT,
            product_name_greenline TEXT,
            confidence_score REAL DEFAULT 0.0,
            match_type TEXT,
            linked_date TIMESTAMP,
            FOREIGN KEY (gov_id) REFERENCES master_catalog(gov_id)
        );
        
        CREATE TABLE IF NOT EXISTS sales_history (
            sale_id INTEGER PRIMARY KEY,
            gov_id TEXT,
            greenline_sku TEXT,
            location TEXT,
            sale_date DATE,
            quantity_sold REAL,
            revenue REAL,
            ingested_timestamp TIMESTAMP,
            FOREIGN KEY (gov_id) REFERENCES master_catalog(gov_id)
        );
        
        CREATE TABLE IF NOT EXISTS inventory_snapshot (
            snapshot_id INTEGER PRIMARY KEY,
            gov_id TEXT,
            greenline_sku TEXT,
            location TEXT,
            snapshot_date DATE,
            on_hand INTEGER,
            ingested_timestamp TIMESTAMP,
            FOREIGN KEY (gov_id) REFERENCES master_catalog(gov_id)
        );
        
        CREATE TABLE IF NOT EXISTS data_lineage (
            lineage_id INTEGER PRIMARY KEY,
            source_file TEXT,
            file_hash TEXT UNIQUE,
            import_timestamp TIMESTAMP,
            gov_ids_touched TEXT,
            rows_processed INTEGER,
            rows_inserted INTEGER,
            rows_updated INTEGER,
            status TEXT
        );
        
        CREATE INDEX IF NOT EXISTS idx_gov_id ON master_catalog(gov_id);
        CREATE INDEX IF NOT EXISTS idx_sku ON sku_linkage(greenline_sku);
        CREATE INDEX IF NOT EXISTS idx_sales_date ON sales_history(sale_date);
        """
        
        self.db.executescript(schema)
        self.db.commit()
    
    # ========== FILE DETECTION ==========
    
    def ingest_file(self, filepath: str) -> Dict:
        """Auto-detect and ingest file"""
        file_type = self._detect_file_type(filepath)
        logger.info(f"Detected file type: {file_type}")
        
        handlers = {
            'government': self._ingest_government_catalog,
            'inventory': self._ingest_inventory,
            'sales': self._ingest_sales,
            'po': self._ingest_po,
        }
        
        handler = handlers.get(file_type)
        if not handler:
            return self._error_response("Unknown file type")
        
        try:
            return handler(filepath)
        except Exception as e:
            logger.error(f"Ingestion error: {e}")
            return self._error_response(str(e))
    
    def _detect_file_type(self, filepath: str) -> str:
        """Smart file type detection"""
        filename = Path(filepath).name.lower()
        
        # Pattern matching
        patterns = {
            'government': ['gov', 'government', 'order', 'catalog'],
            'inventory': ['inventory', 'stock', 'warehouse'],
            'sales': ['sales', 'revenue', 'transaction'],
            'po': ['po', 'purchase', 'order'],
        }
        
        for file_type, keywords in patterns.items():
            if any(kw in filename for kw in keywords):
                return file_type
        
        # Check Excel sheet names
        if filename.endswith(('.xlsx', '.xlsm')):
            try:
                xl = pd.ExcelFile(filepath)
                sheets = [s.lower() for s in xl.sheet_names]
                for file_type, keywords in patterns.items():
                    if any(kw in ' '.join(sheets) for kw in keywords):
                        return file_type
            except:
                pass
        
        # Check CSV columns
        if filename.endswith('.csv'):
            try:
                df = pd.read_csv(filepath, nrows=0)
                cols = ' '.join([c.lower() for c in df.columns])
                for file_type, keywords in patterns.items():
                    if any(kw in cols for kw in keywords):
                        return file_type
            except:
                pass
        
        return 'unknown'
    
    # ========== GOVERNMENT CATALOG ==========
    
    def _ingest_government_catalog(self, filepath: str) -> Dict:
        """Parse official government product list"""
        try:
            df = pd.read_excel(filepath) if filepath.endswith(('.xlsx', '.xlsm')) else pd.read_csv(filepath)
            df.columns = [c.strip().lower() for c in df.columns]
            
            required = ['gov_id', 'product_name', 'wholesale_cost']
            if not all(col in df.columns for col in required):
                return self._error_response(f"Missing columns: {required}")
            
            new_count = 0
            updated_count = 0
            govs = []
            
            for _, row in df.iterrows():
                gov_id = str(row['gov_id']).strip()
                if not gov_id or gov_id == 'nan':
                    continue
                
                govs.append(gov_id)
                
                self.cursor.execute("SELECT 1 FROM master_catalog WHERE gov_id = ?", (gov_id,))
                exists = self.cursor.fetchone()
                
                if exists:
                    self.cursor.execute("""
                        UPDATE master_catalog
                        SET product_name = ?, wholesale_cost = ?, last_updated = ?
                        WHERE gov_id = ?
                    """, (row.get('product_name', ''), float(row.get('wholesale_cost', 0)), datetime.now(), gov_id))
                    updated_count += 1
                else:
                    self.cursor.execute("""
                        INSERT INTO master_catalog
                        (gov_id, product_name, wholesale_cost, case_size, last_updated, data_source)
                        VALUES (?, ?, ?, ?, ?, ?)
                    """, (gov_id, row.get('product_name', ''), float(row.get('wholesale_cost', 0)),
                          int(row.get('case_size', 1)), datetime.now(), 'government'))
                    new_count += 1
            
            self.db.commit()
            self._record_lineage(filepath, govs, new_count, updated_count, 'success')
            
            return {
                'success': True,
                'message': f"✅ KNOWLEDGE BASE UPDATED\n  • New Products: {new_count}\n  • Updated: {updated_count}",
                'govs_touched': govs,
                'new_products': new_count,
                'updated_products': updated_count,
                'confidence_score': 1.0,
                'action': 'knowledge_base_updated'
            }
        except Exception as e:
            return self._error_response(str(e))
    
    # ========== INVENTORY ==========
    
    def _ingest_inventory(self, filepath: str) -> Dict:
        """Parse and link inventory to master catalog"""
        try:
            df = pd.read_csv(filepath) if filepath.endswith('.csv') else pd.read_excel(filepath)
            df.columns = [c.strip().lower() for c in df.columns]
            
            sku_col = next((c for c in df.columns if 'sku' in c), None)
            qty_col = next((c for c in df.columns if any(x in c for x in ['qty', 'on_hand', 'stock'])), None)
            name_col = next((c for c in df.columns if any(x in c for x in ['product', 'name'])), None)
            
            if not all([sku_col, qty_col]):
                return self._error_response("Missing SKU or Quantity columns")
            
            linked = 0
            govs = set()
            
            for _, row in df.iterrows():
                sku = str(row[sku_col]).strip()
                qty = float(row[qty_col])
                name = str(row[name_col]) if name_col else ""
                
                gov_id, confidence = self._find_gov_id(sku, name)
                if gov_id:
                    govs.add(gov_id)
                    self._upsert_linkage(sku, gov_id, name, confidence)
                    self._upsert_inventory(gov_id, sku, qty)
                    linked += 1
            
            self.db.commit()
            self._record_lineage(filepath, list(govs), linked, 0, 'success')
            
            avg_conf = self._avg_confidence(list(govs)) if govs else 0.0
            
            return {
                'success': True,
                'message': f"✅ LINKED TO MASTER CATALOG\n  • Matched: {linked} items\n  • Confidence: {avg_conf:.1%}",
                'govs_touched': list(govs),
                'new_products': linked,
                'updated_products': 0,
                'confidence_score': avg_conf,
                'action': 'inventory_linked'
            }
        except Exception as e:
            return self._error_response(str(e))
    
    # ========== SALES ==========
    
    def _ingest_sales(self, filepath: str) -> Dict:
        """Parse and record sales history"""
        try:
            df = pd.read_csv(filepath) if filepath.endswith('.csv') else pd.read_excel(filepath)
            df.columns = [c.strip().lower() for c in df.columns]
            
            sku_col = next((c for c in df.columns if 'sku' in c), None)
            date_col = next((c for c in df.columns if 'date' in c), None)
            qty_col = next((c for c in df.columns if any(x in c for x in ['qty', 'sold'])), None)
            
            if not all([sku_col, date_col, qty_col]):
                return self._error_response("Missing SKU, Date, or Quantity columns")
            
            inserted = 0
            govs = set()
            
            for _, row in df.iterrows():
                sku = str(row[sku_col]).strip()
                sale_date = pd.to_datetime(row[date_col]).date()
                qty = float(row[qty_col])
                
                gov_id = self._lookup_gov_id(sku)
                if gov_id:
                    govs.add(gov_id)
                    self.cursor.execute("""
                        INSERT INTO sales_history
                        (gov_id, greenline_sku, location, sale_date, quantity_sold, ingested_timestamp)
                        VALUES (?, ?, ?, ?, ?, ?)
                    """, (gov_id, sku, row.get('location', 'Default'), sale_date, qty, datetime.now()))
                    inserted += 1
            
            self.db.commit()
            self._record_lineage(filepath, list(govs), inserted, 0, 'success')
            
            return {
                'success': True,
                'message': f"✅ SALES HISTORY INGESTED\n  • Records: {inserted}",
                'govs_touched': list(govs),
                'new_products': inserted,
                'updated_products': 0,
                'confidence_score': 0.95,
                'action': 'sales_ingested'
            }
        except Exception as e:
            return self._error_response(str(e))
    
    # ========== PURCHASE ORDERS ==========
    
    def _ingest_po(self, filepath: str) -> Dict:
        """Parse purchase orders"""
        try:
            df = pd.read_csv(filepath) if filepath.endswith('.csv') else pd.read_excel(filepath)
            df.columns = [c.strip().lower() for c in df.columns]
            
            sku_col = next((c for c in df.columns if 'sku' in c), None)
            if not sku_col:
                return self._error_response("Missing SKU column")
            
            linked = 0
            govs = set()
            
            for _, row in df.iterrows():
                sku = str(row[sku_col]).strip()
                gov_id = self._lookup_gov_id(sku)
                if gov_id:
                    govs.add(gov_id)
                    linked += 1
            
            self._record_lineage(filepath, list(govs), linked, 0, 'success')
            
            return {
                'success': True,
                'message': f"✅ PO LINKED TO MASTER CATALOG\n  • Purchase Orders: {linked}",
                'govs_touched': list(govs),
                'new_products': linked,
                'updated_products': 0,
                'confidence_score': 0.9,
                'action': 'po_linked'
            }
        except Exception as e:
            return self._error_response(str(e))
    
    # ========== LINKAGE ENGINE (Rosetta Stone) ==========
    
    def _find_gov_id(self, sku: str, product_name: str = "") -> Tuple[Optional[str], float]:
        """Find Gov ID for SKU using multiple strategies"""
        # Strategy 1: Exact match
        self.cursor.execute("SELECT gov_id, confidence_score FROM sku_linkage WHERE greenline_sku = ?", (sku,))
        result = self.cursor.fetchone()
        if result:
            return result
        
        # Strategy 2: Fuzzy match on product names
        if product_name:
            self.cursor.execute("SELECT gov_id, product_name FROM master_catalog")
            for gov_id, gov_name in self.cursor.fetchall():
                similarity = self._fuzzy_match(product_name, gov_name)
                if similarity > 0.85:
                    return (gov_id, similarity)
        
        return (None, 0.0)
    
    def _lookup_gov_id(self, sku: str) -> Optional[str]:
        """Quick lookup of Gov ID"""
        self.cursor.execute("SELECT gov_id FROM sku_linkage WHERE greenline_sku = ?", (sku,))
        result = self.cursor.fetchone()
        return result[0] if result else None
    
    def _fuzzy_match(self, s1: str, s2: str) -> float:
        """Fuzzy string matching"""
        return SequenceMatcher(None, s1.lower(), s2.lower()).ratio()
    
    def _upsert_linkage(self, sku: str, gov_id: str, name: str, confidence: float):
        """Store SKU-to-Gov-ID mapping"""
        self.cursor.execute("SELECT linkage_id FROM sku_linkage WHERE greenline_sku = ?", (sku,))
        exists = self.cursor.fetchone()
        
        if exists:
            self.cursor.execute("""
                UPDATE sku_linkage SET gov_id = ?, confidence_score = ?, linked_date = ?
                WHERE greenline_sku = ?
            """, (gov_id, confidence, datetime.now(), sku))
        else:
            self.cursor.execute("""
                INSERT INTO sku_linkage
                (greenline_sku, gov_id, product_name_greenline, confidence_score, match_type, linked_date)
                VALUES (?, ?, ?, ?, ?, ?)
            """, (sku, gov_id, name, confidence, 'fuzzy', datetime.now()))
    
    def _upsert_inventory(self, gov_id: str, sku: str, qty: float):
        """Store inventory snapshot"""
        self.cursor.execute("""
            INSERT INTO inventory_snapshot
            (gov_id, greenline_sku, location, snapshot_date, on_hand, ingested_timestamp)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (gov_id, sku, 'Default', datetime.now().date(), int(qty), datetime.now()))
    
    def _record_lineage(self, filepath: str, govs: List[str], inserted: int, updated: int, status: str):
        """Audit trail"""
        try:
            file_hash = hashlib.md5(open(filepath, 'rb').read()).hexdigest()
            self.cursor.execute("""
                INSERT INTO data_lineage
                (source_file, file_hash, import_timestamp, gov_ids_touched, rows_processed, rows_inserted, rows_updated, status)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (Path(filepath).name, file_hash, datetime.now(), json.dumps(govs), len(govs), inserted, updated, status))
        except sqlite3.IntegrityError:
            logger.warning(f"Duplicate import detected: {filepath}")
    
    def _avg_confidence(self, govs: List[str]) -> float:
        """Calculate average linkage confidence"""
        if not govs:
            return 0.0
        placeholders = ','.join(['?'] * len(govs))
        self.cursor.execute(f"SELECT AVG(confidence_score) FROM sku_linkage WHERE gov_id IN ({placeholders})", govs)
        result = self.cursor.fetchone()[0]
        return result if result else 0.0
    
    def get_status(self) -> Dict:
        """Get knowledge base status"""
        self.cursor.execute("SELECT COUNT(*) FROM master_catalog")
        products = self.cursor.fetchone()[0]
        
        self.cursor.execute("SELECT COUNT(*) FROM sales_history")
        sales_records = self.cursor.fetchone()[0]
        
        self.cursor.execute("SELECT AVG(confidence_score) FROM sku_linkage")
        confidence = self.cursor.fetchone()[0] or 0.0
        
        return {
            'products': products,
            'sales_records': sales_records,
            'confidence': confidence,
            'db_path': self.db_path
        }
    
    def _error_response(self, message: str) -> Dict:
        """Standard error response"""
        return {
            'success': False,
            'message': f"❌ {message}",
            'govs_touched': [],
            'new_products': 0,
            'updated_products': 0,
            'confidence_score': 0.0,
            'action': 'error'
        }


if __name__ == '__main__':
    # Test initialization
    kb = KnowledgeBase()
    print(f"Knowledge Base initialized: {kb.get_status()}")
