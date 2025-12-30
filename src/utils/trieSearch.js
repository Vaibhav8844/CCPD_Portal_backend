/**
 * Trie (Prefix Tree) implementation for ultra-fast autocomplete/search
 * Time Complexity:
 *  - Insert: O(m) where m is the length of the word
 *  - Search: O(m) where m is the length of the prefix
 *  - Space: O(n * m) where n is number of words and m is average word length
 * 
 * This is much faster than includes() which is O(n * m) for each search
 * where n is the number of users and m is the average string length.
 */

class TrieNode {
  constructor() {
    this.children = new Map();
    this.isEndOfWord = false;
    this.data = null; // Store associated data (user object)
  }
}

export class Trie {
  constructor() {
    this.root = new TrieNode();
  }

  /**
   * Insert a word into the Trie with associated data
   * @param {string} word - The word/prefix to insert
   * @param {Object} data - Associated data (e.g., user object)
   */
  insert(word, data) {
    if (!word) return;
    
    word = word.toLowerCase().trim();
    let node = this.root;

    for (const char of word) {
      if (!node.children.has(char)) {
        node.children.set(char, new TrieNode());
      }
      node = node.children.get(char);
    }

    node.isEndOfWord = true;
    node.data = data;
  }

  /**
   * Search for all words that start with the given prefix
   * @param {string} prefix - The prefix to search for
   * @returns {Array} Array of data objects that match the prefix
   */
  searchPrefix(prefix) {
    if (!prefix) return [];
    
    prefix = prefix.toLowerCase().trim();
    let node = this.root;

    // Navigate to the prefix node
    for (const char of prefix) {
      if (!node.children.has(char)) {
        return []; // Prefix not found
      }
      node = node.children.get(char);
    }

    // Collect all words that start with this prefix
    const results = [];
    this._collectAllWords(node, results);
    return results;
  }

  /**
   * Recursively collect all words from a given node
   * @param {TrieNode} node - Current node
   * @param {Array} results - Results array to populate
   * @private
   */
  _collectAllWords(node, results) {
    if (node.isEndOfWord && node.data) {
      results.push(node.data);
    }

    for (const child of node.children.values()) {
      this._collectAllWords(child, results);
    }
  }

  /**
   * Check if a word exists in the Trie
   * @param {string} word - The word to search for
   * @returns {boolean}
   */
  search(word) {
    if (!word) return false;
    
    word = word.toLowerCase().trim();
    let node = this.root;

    for (const char of word) {
      if (!node.children.has(char)) {
        return false;
      }
      node = node.children.get(char);
    }

    return node.isEndOfWord;
  }

  /**
   * Clear the entire Trie
   */
  clear() {
    this.root = new TrieNode();
  }
}

/**
 * Build a Trie index for user search by name and email
 * @param {Array} users - Array of user objects with name and email
 * @returns {Trie} Populated Trie instance
 */
export function buildUserSearchIndex(users) {
  const trie = new Trie();

  for (const user of users) {
    // Index by name (full name)
    if (user.name) {
      trie.insert(user.name, user);
      
      // Also index individual words in the name (for partial matching)
      const words = user.name.toLowerCase().split(/\s+/);
      for (const word of words) {
        if (word.length > 0) {
          trie.insert(word, user);
        }
      }
    }

    // Index by email
    if (user.email) {
      trie.insert(user.email, user);
      
      // Index email username part (before @)
      const emailParts = user.email.split('@');
      if (emailParts[0]) {
        trie.insert(emailParts[0], user);
      }
    }
  }

  return trie;
}

/**
 * Deduplicate search results (since same user can match multiple indexed terms)
 * @param {Array} results - Array of user objects
 * @returns {Array} Deduplicated array based on email
 */
export function deduplicateResults(results) {
  const seen = new Set();
  const unique = [];

  for (const user of results) {
    if (!seen.has(user.email)) {
      seen.add(user.email);
      unique.push(user);
    }
  }

  return unique;
}

/**
 * Fuzzy search implementation (optional enhancement)
 * Allows 1 character difference for better UX
 * @param {string} str1 
 * @param {string} str2 
 * @returns {number} Levenshtein distance
 */
export function levenshteinDistance(str1, str2) {
  const m = str1.length;
  const n = str2.length;
  const dp = Array(m + 1).fill(null).map(() => Array(n + 1).fill(0));

  for (let i = 0; i <= m; i++) dp[i][0] = i;
  for (let j = 0; j <= n; j++) dp[0][j] = j;

  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      if (str1[i - 1] === str2[j - 1]) {
        dp[i][j] = dp[i - 1][j - 1];
      } else {
        dp[i][j] = 1 + Math.min(
          dp[i - 1][j],     // deletion
          dp[i][j - 1],     // insertion
          dp[i - 1][j - 1]  // substitution
        );
      }
    }
  }

  return dp[m][n];
}
