import unittest
import tkinter as tk
from main import CigarInventory
import json
import os
from datetime import datetime
import shutil

class TestCigarInventory(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        """Set up test environment once before all tests."""
        # Create test directory
        cls.test_dir = os.path.join(os.path.expanduser('~'), 'Documents', 'CigarInventoryTest')
        os.makedirs(cls.test_dir, exist_ok=True)
        
        # Store original paths
        cls.original_dir = os.path.join(os.path.expanduser('~'), 'Documents', 'CigarInventory')
        cls.original_inventory_path = os.path.join(cls.original_dir, 'cigar_inventory.json')
        cls.original_sales_path = os.path.join(cls.original_dir, 'sales_history.json')
        cls.original_brands_path = os.path.join(cls.original_dir, 'cigar_brands.json')
        cls.original_sizes_path = os.path.join(cls.original_dir, 'cigar_sizes.json')
        cls.original_types_path = os.path.join(cls.original_dir, 'cigar_types.json')

    def setUp(self):
        """Set up test environment before each test."""
        self.root = tk.Tk()
        self.app = CigarInventory(self.root)
        
        # Store original paths
        self._original_paths = {
            'inventory': self.app.inventory_file,
            'sales': self.app.sales_file,
            'brands': self.app.brands_file,
            'sizes': self.app.sizes_file,
            'types': self.app.types_file
        }
        
        # Update paths for testing
        self.app.inventory_file = os.path.join(self.test_dir, 'cigar_inventory.json')
        self.app.sales_file = os.path.join(self.test_dir, 'sales_history.json')
        self.app.brands_file = os.path.join(self.test_dir, 'cigar_brands.json')
        self.app.sizes_file = os.path.join(self.test_dir, 'cigar_sizes.json')
        self.app.types_file = os.path.join(self.test_dir, 'cigar_types.json')
        
        # Clear sales history at start of each test
        self.app.sales_history = []
        
        # Sample test data
        self.test_cigar = {
            'brand': 'Test Brand',
            'cigar': 'Test Cigar',
            'size': 'Robusto',
            'type': 'Regular',
            'count': 10,
            'price': 100.00,
            'shipping': 10.00,
            'price_per_stick': 0.0,
            'personal_rating': 5
        }

    def tearDown(self):
        """Clean up after each test."""
        # Restore original paths
        self.app.inventory_file = self._original_paths['inventory']
        self.app.sales_file = self._original_paths['sales']
        self.app.brands_file = self._original_paths['brands']
        self.app.sizes_file = self._original_paths['sizes']
        self.app.types_file = self._original_paths['types']
        
        # Clean up test files
        for filename in os.listdir(self.test_dir):
            file_path = os.path.join(self.test_dir, filename)
            try:
                if os.path.isfile(file_path):
                    os.unlink(file_path)
            except Exception as e:
                print(f'Error deleting {file_path}: {e}')
                
        self.root.destroy()

    @classmethod
    def tearDownClass(cls):
        """Clean up after all tests are done."""
        # Remove test directory
        try:
            shutil.rmtree(cls.test_dir)
        except Exception as e:
            print(f'Error removing test directory: {e}')

    def test_calculate_price_per_stick(self):
        """Test price per stick calculation."""
        # Test normal calculation
        price_per_stick = self.app.calculate_price_per_stick(100.00, 10.00, 10)
        expected = (100.00 * (1 + self.app.tax_rate) + 10.00) / 10
        self.assertAlmostEqual(price_per_stick, expected, places=2)

        # Test zero count
        self.assertEqual(self.app.calculate_price_per_stick(100.00, 10.00, 0), 0)

        # Test zero price and shipping
        self.assertEqual(self.app.calculate_price_per_stick(0, 0, 10), 0)

    def test_sorting(self):
        """Test sorting functionality."""
        # Reset sort state
        self.app.sort_reverse = {}
        self.app.sort_column = None
        
        # Clear existing inventory and add test data
        self.app.inventory = [
            {
                'brand': 'Padron',
                'cigar': 'Serie 1926',
                'count': 5,
                'price': 20.00,
            },
            {
                'brand': 'Arturo Fuente',
                'cigar': 'Opus X',
                'count': 10,
                'price': 30.00,
            },
            {
                'brand': 'Cohiba',
                'cigar': 'Behike',
                'count': 15,
                'price': 25.00,
            }
        ]

        # Test brand sorting (ascending)
        print("\nTesting brand sort (ascending):")
        print("Before sort:", [cigar['brand'] for cigar in self.app.inventory])
        self.app.sort_reverse['brand'] = False  # Force ascending sort
        self.app.sort_treeview('brand')
        sorted_brands = [cigar['brand'] for cigar in self.app.inventory]
        print("After sort:", sorted_brands)
        expected_brands = ['Arturo Fuente', 'Cohiba', 'Padron']
        self.assertEqual(sorted_brands, expected_brands, 
                        f"Expected brands to be {expected_brands}, but got {sorted_brands}")

        # Test price sorting (ascending)
        print("\nTesting price sort (ascending):")
        print("Before sort:", [cigar['price'] for cigar in self.app.inventory])
        self.app.sort_reverse['price'] = False  # Force ascending sort
        self.app.sort_treeview('price')
        sorted_prices = [cigar['price'] for cigar in self.app.inventory]
        print("After sort:", sorted_prices)
        expected_prices = [20.00, 25.00, 30.00]
        self.assertEqual(sorted_prices, expected_prices,
                        f"Expected prices to be {expected_prices}, but got {sorted_prices}")

        # Test count sorting (ascending)
        print("\nTesting count sort (ascending):")
        print("Before sort:", [cigar['count'] for cigar in self.app.inventory])
        self.app.sort_reverse['count'] = False  # Force ascending sort
        self.app.sort_treeview('count')
        sorted_counts = [cigar['count'] for cigar in self.app.inventory]
        print("After sort:", sorted_counts)
        expected_counts = [5, 10, 15]
        self.assertEqual(sorted_counts, expected_counts,
                        f"Expected counts to be {expected_counts}, but got {sorted_counts}")

    def test_inventory_totals(self):
        """Test inventory totals calculation."""
        # Add test data
        self.app.inventory = [
            {'count': 10, 'price_per_stick': 10.00, 'shipping': 5.00},
            {'count': 5, 'price_per_stick': 20.00, 'shipping': 10.00}
        ]

        self.app.update_inventory_totals()
        
        # Test total count
        total_count = sum(int(cigar['count']) for cigar in self.app.inventory)
        self.assertEqual(total_count, 15)
        
        # Test total value
        expected_value = (10 * 10.00) + (5 * 20.00)
        actual_value = sum(int(cigar['count']) * float(cigar['price_per_stick']) 
                          for cigar in self.app.inventory)
        self.assertAlmostEqual(actual_value, expected_value, places=2)

    def test_sale_processing(self):
        """Test sale processing functionality."""
        # Add test inventory
        self.app.inventory = [self.test_cigar]
        initial_count = self.test_cigar['count']
        
        # Simulate sale
        self.app.checkbox_states = {self.test_cigar['cigar']: True}
        self.app.quantity_spinboxes = {
            self.test_cigar['cigar']: tk.StringVar(value='2')
        }
        
        # Process sale
        self.app.sell_selected()
        
        # Verify inventory count decreased
        self.assertEqual(self.test_cigar['count'], initial_count - 2)
        
        # Verify sale record was created
        self.assertEqual(len(self.app.sales_history), 1)
        sale_record = self.app.sales_history[0]
        self.assertEqual(sale_record['quantity'], 2)
        self.assertEqual(sale_record['cigar'], self.test_cigar['cigar'])

    def test_shipping_calculator(self):
        """Test shipping calculator functionality."""
        # Test shipping cost calculation
        shipping_cost = 10.00
        total_cigars = 5
        
        self.app.ship_cost.insert(0, str(shipping_cost))
        self.app.total_cigars.insert(0, str(total_cigars))
        self.app.calculate_shipping()
        
        expected_per_stick = shipping_cost / total_cigars
        expected_five_pack = shipping_cost / 5
        expected_ten_pack = shipping_cost / 10
        
        # Verify calculations
        self.assertAlmostEqual(float(self.app.per_stick.cget('text').split('$')[1].split(')')[0]), 
                              expected_per_stick, places=2)
        self.assertAlmostEqual(float(self.app.five_pack.cget('text').split('$')[1]), 
                              expected_five_pack, places=2)
        self.assertAlmostEqual(float(self.app.ten_pack.cget('text').split('$')[1]), 
                              expected_ten_pack, places=2)

    def test_data_persistence(self):
        """Test saving and loading of inventory data."""
        # Add test data
        self.app.inventory = [self.test_cigar]
        
        # Save inventory
        self.app.save_inventory()
        
        # Clear inventory
        self.app.inventory = []
        
        # Load inventory
        self.app.load_inventory()
        
        # Verify data was preserved
        self.assertEqual(len(self.app.inventory), 1)
        loaded_cigar = self.app.inventory[0]
        self.assertEqual(loaded_cigar['brand'], self.test_cigar['brand'])
        self.assertEqual(loaded_cigar['count'], self.test_cigar['count'])

    def test_search_functionality(self):
        """Test search functionality."""
        # Add test data
        self.app.inventory = [
            {'brand': 'Test Brand 1', 'cigar': 'Test Cigar 1'},
            {'brand': 'Test Brand 2', 'cigar': 'Test Cigar 2'},
            {'brand': 'Other Brand', 'cigar': 'Other Cigar'}
        ]
        
        # Test search
        self.app.search_var.set('Test')
        self.app.refresh_inventory()
        
        # Count visible items that match search
        visible_items = len(self.app.tree.get_children())
        self.assertEqual(visible_items, 2)

if __name__ == '__main__':
    unittest.main()