//! Generic LRU (Least Recently Used) cache.
//!
//! Provides a simple LRU eviction cache used by text measurement,
//! text wrapping, and other render subsystems.

use std::collections::{HashMap, VecDeque};
use std::hash::Hash;

/// A simple LRU cache with a fixed capacity.
///
/// When the cache exceeds capacity, the oldest (least recently inserted)
/// entries are evicted. This is an insertion-order LRU — lookups do not
/// promote entries.
pub struct LruCache<K: Hash + Eq + Clone, V> {
    entries: HashMap<K, V>,
    order: VecDeque<K>,
    capacity: usize,
}

impl<K: Hash + Eq + Clone, V> LruCache<K, V> {
    /// Create a new cache with the given capacity.
    ///
    /// A capacity of 0 disables caching entirely.
    pub fn new(capacity: usize) -> Self {
        Self {
            entries: HashMap::new(),
            order: VecDeque::new(),
            capacity,
        }
    }

    /// Look up a value by key. Returns `None` if not present or capacity is 0.
    pub fn get(&self, key: &K) -> Option<&V> {
        if self.capacity == 0 {
            return None;
        }
        self.entries.get(key)
    }

    /// Insert a key-value pair. If the key already exists, the value is NOT updated.
    /// Returns `true` if the entry was newly inserted, `false` if it already existed.
    pub fn insert(&mut self, key: K, value: V) -> bool {
        if self.capacity == 0 {
            return false;
        }
        if self.entries.contains_key(&key) {
            return false;
        }
        self.entries.insert(key.clone(), value);
        self.order.push_back(key);
        self.enforce_cap();
        true
    }

    /// Insert a value if not present, then return a reference.
    ///
    /// If the key already exists, the existing value is returned unchanged.
    /// If the key is new, the value is inserted and a reference returned.
    /// Returns `None` only if capacity is 0.
    pub fn get_or_insert(&mut self, key: &K, value: V) -> Option<&V> {
        if self.capacity == 0 {
            return None;
        }
        if !self.entries.contains_key(key) {
            self.entries.insert(key.clone(), value);
            self.order.push_back(key.clone());
            self.enforce_cap();
        }
        self.entries.get(key)
    }

    /// Check if a key is present.
    pub fn contains_key(&self, key: &K) -> bool {
        self.entries.contains_key(key)
    }

    /// Number of cached entries.
    pub fn len(&self) -> usize {
        self.entries.len()
    }

    /// Whether the cache is empty.
    pub fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }

    /// Remove all entries.
    pub fn clear(&mut self) {
        self.entries.clear();
        self.order.clear();
    }

    /// Evict oldest entries until we're at or below capacity.
    fn enforce_cap(&mut self) {
        while self.entries.len() > self.capacity {
            if let Some(oldest) = self.order.pop_front() {
                self.entries.remove(&oldest);
            } else {
                break;
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_basic_insert_get() {
        let mut cache: LruCache<String, i32> = LruCache::new(3);
        cache.insert("a".to_string(), 1);
        cache.insert("b".to_string(), 2);
        cache.insert("c".to_string(), 3);

        assert_eq!(cache.get(&"a".to_string()), Some(&1));
        assert_eq!(cache.get(&"b".to_string()), Some(&2));
        assert_eq!(cache.get(&"c".to_string()), Some(&3));
        assert_eq!(cache.len(), 3);
    }

    #[test]
    fn test_eviction() {
        let mut cache: LruCache<String, i32> = LruCache::new(2);
        cache.insert("a".to_string(), 1);
        cache.insert("b".to_string(), 2);
        cache.insert("c".to_string(), 3);

        // "a" should be evicted
        assert_eq!(cache.get(&"a".to_string()), None);
        assert_eq!(cache.get(&"b".to_string()), Some(&2));
        assert_eq!(cache.get(&"c".to_string()), Some(&3));
        assert_eq!(cache.len(), 2);
    }

    #[test]
    fn test_zero_capacity() {
        let mut cache: LruCache<String, i32> = LruCache::new(0);
        assert!(!cache.insert("a".to_string(), 1));
        assert_eq!(cache.get(&"a".to_string()), None);
        assert!(cache.is_empty());
    }

    #[test]
    fn test_duplicate_insert() {
        let mut cache: LruCache<String, i32> = LruCache::new(3);
        assert!(cache.insert("a".to_string(), 1));
        assert!(!cache.insert("a".to_string(), 2)); // should not update
        assert_eq!(cache.get(&"a".to_string()), Some(&1)); // original value
    }

    #[test]
    fn test_clear() {
        let mut cache: LruCache<String, i32> = LruCache::new(3);
        cache.insert("a".to_string(), 1);
        cache.insert("b".to_string(), 2);
        cache.clear();
        assert!(cache.is_empty());
        assert_eq!(cache.get(&"a".to_string()), None);
    }
}
