/**
 * String formatting options
 */
export interface StringFormatOptions {
  /** Maximum length */
  maxLength?: number;
  /** Truncation indicator */
  truncationIndicator?: string;
  /** Case transformation */
  case?: 'upper' | 'lower' | 'title' | 'sentence' | 'camel' | 'pascal' | 'snake' | 'kebab';
  /** Padding configuration */
  padding?: {
    /** Total width */
    width: number;
    /** Pad character */
    char?: string;
    /** Pad side */
    side?: 'left' | 'right' | 'both';
  };
  /** Trim whitespace */
  trim?: boolean;
  /** Replace patterns */
  replace?: Array<{
    pattern: RegExp | string;
    replacement: string;
  }>;
  /** Prefix to add */
  prefix?: string;
  /** Suffix to add */
  suffix?: string;
  /** Escape special characters */
  escape?: boolean;
  /** Remove accents/diacritics */
  normalize?: boolean;
  /** Format as currency */
  currency?: {
    symbol?: string;
    position?: 'before' | 'after';
    thousandsSeparator?: boolean;
    decimalPlaces?: number;
  };
  /** Format as phone number */
  phone?: {
    format: 'international' | 'national' | 'e164';
    countryCode?: string;
    separator?: string;
  };
  /** Format as credit card */
  creditCard?: {
    mask?: boolean;
    separator?: 'space' | 'dash' | 'none';
    maskChar?: string;
  };
  /** Format as email */
  email?: {
    lowercase?: boolean;
    trim?: boolean;
  };
  /** Format as URL */
  url?: {
    protocol?: boolean;
    www?: boolean;
    lowercase?: boolean;
  };
  /** Format as file path */
  filePath?: {
    separator?: 'forward' | 'backward';
    normalize?: boolean;
    extension?: string;
  };
}

/**
 * String manipulation utilities
 */
export class StringFormatter {
  /**
   * Format a string with options
   */
  static format(value: any, options?: StringFormatOptions): string {
    if (value === null || value === undefined) {
      return '';
    }

    let str = String(value);

    if (!options) {
      return str;
    }

    // Apply transformations in order
    if (options.trim) {
      str = str.trim();
    }

    if (options.normalize) {
      str = this.normalize(str);
    }

    if (options.replace) {
      options.replace.forEach(({ pattern, replacement }) => {
        str = str.replace(new RegExp(pattern), replacement);
      });
    }

    if (options.case) {
      str = this.transformCase(str, options.case);
    }

    if (options.currency) {
      str = this.formatAsCurrency(str, options.currency);
    }

    if (options.phone) {
      str = this.formatAsPhone(str, options.phone);
    }

    if (options.creditCard) {
      str = this.formatAsCreditCard(str, options.creditCard);
    }

    if (options.email) {
      str = this.formatAsEmail(str, options.email);
    }

    if (options.url) {
      str = this.formatAsUrl(str, options.url);
    }

    if (options.filePath) {
      str = this.formatAsFilePath(str, options.filePath);
    }

    if (options.escape) {
      str = this.escapeHtml(str);
    }

    if (options.prefix && !str.startsWith(options.prefix)) {
      str = options.prefix + str;
    }

    if (options.suffix && !str.endsWith(options.suffix)) {
      str = str + options.suffix;
    }

    if (options.padding) {
      str = this.pad(str, options.padding);
    }

    if (options.maxLength && str.length > options.maxLength) {
      const indicator = options.truncationIndicator || '...';
      str = str.substring(0, options.maxLength - indicator.length) + indicator;
    }

    return str;
  }

  /**
   * Transform string case
   */
  static transformCase(str: string, type: StringFormatOptions['case']): string {
    switch (type) {
      case 'upper':
        return str.toUpperCase();

      case 'lower':
        return str.toLowerCase();

      case 'title':
        return str.replace(/\w\S*/g, (word) => {
          return word.charAt(0).toUpperCase() + word.substr(1).toLowerCase();
        });

      case 'sentence':
        return str.charAt(0).toUpperCase() + str.substr(1).toLowerCase();

      case 'camel':
        return str
          .replace(/[-_\s]+(.)?/g, (_, c) => (c ? c.toUpperCase() : ''))
          .replace(/^[A-Z]/, (c) => c.toLowerCase());

      case 'pascal':
        return str
          .replace(/[-_\s]+(.)?/g, (_, c) => (c ? c.toUpperCase() : ''))
          .replace(/^[a-z]/, (c) => c.toUpperCase());

      case 'snake':
        return str
          .replace(/\s+/g, '_')
          .replace(/([a-z])([A-Z])/g, '$1_$2')
          .toLowerCase();

      case 'kebab':
        return str
          .replace(/\s+/g, '-')
          .replace(/([a-z])([A-Z])/g, '$1-$2')
          .toLowerCase();

      default:
        return str;
    }
  }

  /**
   * Pad a string to a specific width
   */
  static pad(str: string, options: NonNullable<StringFormatOptions['padding']>): string {
    const { width, char = ' ', side = 'left' } = options;
    
    if (str.length >= width) {
      return str;
    }

    const padding = char.repeat(width - str.length);

    switch (side) {
      case 'left':
        return padding + str;
      case 'right':
        return str + padding;
      case 'both':
        const leftPad = Math.floor(padding.length / 2);
        const rightPad = padding.length - leftPad;
        return char.repeat(leftPad) + str + char.repeat(rightPad);
      default:
        return str;
    }
  }

  /**
   * Truncate string to maximum length
   */
  static truncate(str: string, maxLength: number, indicator: string = '...'): string {
    if (str.length <= maxLength) {
      return str;
    }
    return str.substring(0, maxLength - indicator.length) + indicator;
  }

  /**
   * Escape HTML special characters
   */
  static escapeHtml(str: string): string {
    const htmlEscapes: Record<string, string> = {
      '&': '&amp;',
      '<': '&lt;',
      '>': '&gt;',
      '"': '&quot;',
      "'": '&#39;',
      '`': '&#96;',
      '=': '&#61;',
    };
    return str.replace(/[&<>"'`=]/g, (char) => htmlEscapes[char] || char);
  }

  /**
   * Unescape HTML special characters
   */
  static unescapeHtml(str: string): string {
    const htmlUnescapes: Record<string, string> = {
      '&amp;': '&',
      '&lt;': '<',
      '&gt;': '>',
      '&quot;': '"',
      '&#39;': "'",
      '&#96;': '`',
      '&#61;': '=',
    };
    return str.replace(/&(?:amp|lt|gt|quot|#39|#96|#61);/g, (entity) => htmlUnescapes[entity] || entity);
  }

  /**
   * Escape regex special characters
   */
  static escapeRegex(str: string): string {
    return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }

  /**
   * Escape CSV special characters
   */
  static escapeCsv(str: string, delimiter: string = ','): string {
    if (str.includes(delimiter) || str.includes('"') || str.includes('\n')) {
      return `"${str.replace(/"/g, '""')}"`;
    }
    return str;
  }

  /**
   * Remove accents/diacritics
   */
  static normalize(str: string): string {
    return str.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  }

  /**
   * Reverse a string
   */
  static reverse(str: string): string {
    return str.split('').reverse().join('');
  }

  /**
   * Check if string contains only ASCII characters
   */
  static isAscii(str: string): boolean {
    return /^[\x00-\x7F]*$/.test(str);
  }

  /**
   * Convert to ASCII (remove non-ASCII)
   */
  static toAscii(str: string): string {
    return str.replace(/[^\x00-\x7F]/g, '');
  }

  /**
   * Extract numbers from string
   */
  static extractNumbers(str: string): string {
    return str.replace(/\D/g, '');
  }

  /**
   * Extract letters from string
   */
  static extractLetters(str: string): string {
    return str.replace(/[^a-zA-Z]/g, '');
  }

  /**
   * Extract alphanumeric characters
   */
  static extractAlphanumeric(str: string): string {
    return str.replace(/[^a-zA-Z0-9]/g, '');
  }

  /**
   * Format as currency
   */
  static formatAsCurrency(str: string, options: NonNullable<StringFormatOptions['currency']>): string {
    const num = parseFloat(str.replace(/[^\d.-]/g, ''));
    if (isNaN(num)) {
      return str;
    }

    const {
      symbol = '$',
      position = 'before',
      thousandsSeparator = true,
      decimalPlaces = 2,
    } = options;

    let formatted = num.toFixed(decimalPlaces);
    
    if (thousandsSeparator) {
      const parts = formatted.split('.');
      parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
      formatted = parts.join('.');
    }

    return position === 'before' 
      ? `${symbol}${formatted}`
      : `${formatted}${symbol}`;
  }

  /**
   * Format as phone number
   */
  static formatAsPhone(str: string, options: NonNullable<StringFormatOptions['phone']>): string {
    const digits = this.extractNumbers(str);
    const { format, countryCode = '1', separator = '-' } = options;

    switch (format) {
      case 'international':
        if (digits.length === 10) {
          return `+${countryCode} ${digits.slice(0, 3)}${separator}${digits.slice(3, 6)}${separator}${digits.slice(6)}`;
        } else if (digits.length === 11 && digits[0] === '1') {
          return `+${digits.slice(0, 1)} ${digits.slice(1, 4)}${separator}${digits.slice(4, 7)}${separator}${digits.slice(7)}`;
        }
        break;

      case 'national':
        if (digits.length === 10) {
          return `(${digits.slice(0, 3)}) ${digits.slice(3, 6)}-${digits.slice(6)}`;
        } else if (digits.length === 11 && digits[0] === '1') {
          return `${digits.slice(0, 1)} (${digits.slice(1, 4)}) ${digits.slice(4, 7)}-${digits.slice(7)}`;
        }
        break;

      case 'e164':
        if (digits.length === 10) {
          return `+${countryCode}${digits}`;
        } else if (digits.length === 11 && digits[0] === '1') {
          return `+${digits}`;
        }
        break;
    }

    return str;
  }

  /**
   * Format as credit card
   */
  static formatAsCreditCard(str: string, options: NonNullable<StringFormatOptions['creditCard']>): string {
    const digits = this.extractNumbers(str);
    const { mask = false, separator = 'space', maskChar = '•' } = options;
    
    let result = digits;
    
    if (mask && result.length > 4) {
      const lastFour = result.slice(-4);
      const masked = result.slice(0, -4).replace(/\d/g, maskChar);
      result = masked + lastFour;
    }

    // Add separators
    const sep = separator === 'space' ? ' ' : separator === 'dash' ? '-' : '';
    if (sep) {
      const groups = [];
      for (let i = 0; i < result.length; i += 4) {
        groups.push(result.slice(i, i + 4));
      }
      result = groups.join(sep);
    }

    return result;
  }

  /**
   * Format as email
   */
  static formatAsEmail(str: string, options: NonNullable<StringFormatOptions['email']>): string {
    let email = str.trim();
    
    if (options.lowercase) {
      email = email.toLowerCase();
    }
    
    if (options.trim && email.includes('@')) {
      const [local, domain] = email.split('@');
      email = `${local}@${domain.toLowerCase()}`;
    }
    
    return email;
  }

  /**
   * Format as URL
   */
  static formatAsUrl(str: string, options: NonNullable<StringFormatOptions['url']>): string {
    let url = str.trim();
    
    if (options.lowercase) {
      url = url.toLowerCase();
    }
    
    if (options.protocol && !url.match(/^[a-zA-Z]+:\/\//)) {
      url = 'https://' + url;
    }
    
    if (options.www && !url.includes('www.') && url.includes('://')) {
      const [protocol, rest] = url.split('://');
      url = `${protocol}://www.${rest}`;
    }
    
    return url;
  }

  /**
   * Format as file path
   */
  static formatAsFilePath(str: string, options: NonNullable<StringFormatOptions['filePath']>): string {
    let path = str.trim();
    
    if (options.normalize) {
      path = path.replace(/[/\\]+/g, options.separator === 'forward' ? '/' : '\\');
    }
    
    if (options.extension && !path.includes('.')) {
      path += '.' + options.extension.replace(/^\./, '');
    }
    
    return path;
  }

  /**
   * Pluralize a word
   */
  static pluralize(word: string, count: number, plural?: string): string {
    if (count === 1) {
      return word;
    }
    
    if (plural) {
      return plural;
    }
    
    // Basic English pluralization rules
    if (word.match(/(s|x|z|ch|sh)$/i)) {
      return word + 'es';
    } else if (word.match(/[^aeiou]y$/i)) {
      return word.slice(0, -1) + 'ies';
    } else if (word.match(/f$/i)) {
      return word.slice(0, -1) + 'ves';
    } else if (word.match(/fe$/i)) {
      return word.slice(0, -2) + 'ves';
    } else {
      return word + 's';
    }
  }

  /**
   * Convert to slug (URL-friendly)
   */
  static slugify(str: string, separator: string = '-'): string {
    return str
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, separator)
      .replace(new RegExp(`^${separator}|${separator}$`, 'g'), '');
  }

  /**
   * Convert to camelCase
   */
  static camelCase(str: string): string {
    return this.transformCase(str, 'camel');
  }

  /**
   * Convert to PascalCase
   */
  static pascalCase(str: string): string {
    return this.transformCase(str, 'pascal');
  }

  /**
   * Convert to snake_case
   */
  static snakeCase(str: string): string {
    return this.transformCase(str, 'snake');
  }

  /**
   * Convert to kebab-case
   */
  static kebabCase(str: string): string {
    return this.transformCase(str, 'kebab');
  }

  /**
   * Capitalize first letter
   */
  static capitalize(str: string): string {
    if (!str) return str;
    return str.charAt(0).toUpperCase() + str.slice(1).toLowerCase();
  }

  /**
   * Capitalize each word
   */
  static capitalizeWords(str: string): string {
    return str.replace(/\b\w/g, (char) => char.toUpperCase());
  }

  /**
   * Check if string starts with any of the provided prefixes
   */
  static startsWithAny(str: string, prefixes: string[]): boolean {
    return prefixes.some(prefix => str.startsWith(prefix));
  }

  /**
   * Check if string ends with any of the provided suffixes
   */
  static endsWithAny(str: string, suffixes: string[]): boolean {
    return suffixes.some(suffix => str.endsWith(suffix));
  }

  /**
   * Remove all whitespace
   */
  static removeWhitespace(str: string): string {
    return str.replace(/\s+/g, '');
  }

  /**
   * Normalize whitespace (replace multiple spaces with single)
   */
  static normalizeWhitespace(str: string): string {
    return str.replace(/\s+/g, ' ').trim();
  }

  /**
   * Get word count
   */
  static wordCount(str: string): number {
    return str.trim().split(/\s+/).length;
  }

  /**
   * Get character count (excluding optional spaces)
   */
  static charCount(str: string, excludeSpaces: boolean = false): number {
    if (excludeSpaces) {
      return str.replace(/\s/g, '').length;
    }
    return str.length;
  }

  /**
   * Mask a string (show only first and last characters)
   */
  static mask(str: string, visibleStart: number = 1, visibleEnd: number = 1, maskChar: string = '*'): string {
    if (str.length <= visibleStart + visibleEnd) {
      return str;
    }
    
    const start = str.slice(0, visibleStart);
    const end = str.slice(-visibleEnd);
    const masked = maskChar.repeat(str.length - visibleStart - visibleEnd);
    
    return start + masked + end;
  }

  /**
   * Generate random string
   */
  static random(length: number, charset: string = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789'): string {
    let result = '';
    for (let i = 0; i < length; i++) {
      result += charset.charAt(Math.floor(Math.random() * charset.length));
    }
    return result;
  }

  /**
   * Check if string is palindrome
   */
  static isPalindrome(str: string): boolean {
    const cleaned = str.toLowerCase().replace(/[^a-z0-9]/g, '');
    return cleaned === cleaned.split('').reverse().join('');
  }

  /**
   * Levenshtein distance between two strings
   */
  static levenshteinDistance(a: string, b: string): number {
    const matrix: number[][] = [];

    for (let i = 0; i <= b.length; i++) {
      matrix[i] = [i];
    }

    for (let j = 0; j <= a.length; j++) {
      matrix[0][j] = j;
    }

    for (let i = 1; i <= b.length; i++) {
      for (let j = 1; j <= a.length; j++) {
        if (b.charAt(i - 1) === a.charAt(j - 1)) {
          matrix[i][j] = matrix[i - 1][j - 1];
        } else {
          matrix[i][j] = Math.min(
            matrix[i - 1][j - 1] + 1, // substitution
            matrix[i][j - 1] + 1,     // insertion
            matrix[i - 1][j] + 1      // deletion
          );
        }
      }
    }

    return matrix[b.length][a.length];
  }

  /**
   * Calculate similarity percentage between two strings
   */
  static similarity(a: string, b: string): number {
    const longer = a.length > b.length ? a : b;
    const shorter = a.length > b.length ? b : a;
    
    if (longer.length === 0) {
      return 1.0;
    }
    
    const distance = this.levenshteinDistance(longer, shorter);
    return (longer.length - distance) / longer.length;
  }

  /**
   * Highlight matching text
   */
  static highlight(text: string, search: string, tag: string = 'mark'): string {
    if (!search) return text;
    
    const regex = new RegExp(this.escapeRegex(search), 'gi');
    return text.replace(regex, `<${tag}>$&</${tag}>`);
  }

  /**
   * Indent text
   */
  static indent(str: string, level: number = 1, char: string = '  '): string {
    const indent = char.repeat(level);
    return str.split('\n').map(line => indent + line).join('\n');
  }

  /**
   * Wrap text to specified width
   */
  static wrap(str: string, width: number, indent: string = ''): string[] {
    const words = str.split(' ');
    const lines: string[] = [];
    let currentLine = indent;

    for (const word of words) {
      if ((currentLine.length + word.length + 1) <= width) {
        currentLine += (currentLine === indent ? '' : ' ') + word;
      } else {
        if (currentLine !== indent) {
          lines.push(currentLine);
        }
        currentLine = indent + word;
      }
    }

    if (currentLine !== indent) {
      lines.push(currentLine);
    }

    return lines;
  }
}

/**
 * Template string utilities
 */
export class TemplateFormatter {
  /**
   * Simple template interpolation
   */
  static interpolate(template: string, data: Record<string, any>, pattern: RegExp = /\{\{([^}]+)\}\}/g): string {
    return template.replace(pattern, (match, key) => {
      const value = key.trim().split('.').reduce((obj: any, prop: string) => obj?.[prop], data);
      return value !== undefined ? String(value) : match;
    });
  }

  /**
   * Template with conditions
   */
  static conditional(template: string, data: Record<string, any>): string {
    // Simple if statements
    let result = template.replace(/\{\{#if ([^}]+)\}\}([\s\S]*?)\{\{\/if\}\}/g, (_unused, condition, content) => {
      const value = this.evaluateExpression(condition, data);
      return value ? content : '';
    });

    // If-else statements
    result = result.replace(/\{\{#if ([^}]+)\}\}([\s\S]*?)\{\{#else\}\}([\s\S]*?)\{\{\/if\}\}/g, (_unused, condition, ifContent, elseContent) => {
      const value = this.evaluateExpression(condition, data);
      return value ? ifContent : elseContent;
    });

    return result;
  }

  /**
   * Template with loops
   */
  static loop(template: string, data: Record<string, any>): string {
    return template.replace(/\{\{#each ([^}]+) as ([^}]+)\}\}([\s\S]*?)\{\{\/each\}\}/g, (_unused, arrayPath, itemName, content) => {
      const array = arrayPath.split('.').reduce((obj: any, prop: string) => obj?.[prop], data);
      
      if (!Array.isArray(array)) {
        return '';
      }

      return array.map((item, index) => {
        const context = {
          ...data,
          [itemName]: item,
          index,
          first: index === 0,
          last: index === array.length - 1,
        };
        return this.interpolate(content, context, /\{\{([^}]+)\}\}/g);
      }).join('');
    });
  }

  /**
   * Evaluate expression safely
   */
  private static evaluateExpression(expression: string, data: Record<string, any>): any {
    try {
      // Simple expression evaluation (for demo purposes)
      // In production, use a proper expression evaluator
      const parts = expression.trim().split(/\s+/);
      
      if (parts.length === 1) {
        return parts[0].split('.').reduce((obj: any, prop: string) => obj?.[prop], data);
      }

      // Handle simple comparisons
      const left = parts[0].split('.').reduce((obj: any, prop: string) => obj?.[prop], data);
      const operator = parts[1];
      const right = parts.slice(2).join(' ').replace(/^['"]|['"]$/g, '');

      switch (operator) {
        case '==':
          return left == right;
        case '===':
          return left === right;
        case '!=':
          return left != right;
        case '!==':
          return left !== right;
        case '>':
          return left > right;
        case '>=':
          return left >= right;
        case '<':
          return left < right;
        case '<=':
          return left <= right;
        default:
          return false;
      }
    } catch {
      return false;
    }
  }
}

/**
 * Byte size formatter
 */
export class ByteFormatter {
  static readonly UNITS = ['B', 'KB', 'MB', 'GB', 'TB', 'PB', 'EB', 'ZB', 'YB'];

  /**
   * Format bytes to human readable string
   */
  static format(bytes: number, decimals: number = 2): string {
    if (bytes === 0) return '0 B';

    const k = 1024;
    const i = Math.floor(Math.log(bytes) / Math.log(k));

    return parseFloat((bytes / Math.pow(k, i)).toFixed(decimals)) + ' ' + this.UNITS[i];
  }

  /**
   * Parse human readable size to bytes
   */
  static parse(size: string): number {
    const match = size.trim().match(/^(\d+(?:\.\d+)?)\s*([kmgtpezy]?b)$/i);
    
    if (!match) {
      return 0;
    }

    const value = parseFloat(match[1]);
    const unit = match[2].toUpperCase();
    const index = this.UNITS.findIndex(u => u === unit);

    if (index === -1) {
      return 0;
    }

    return value * Math.pow(1024, index);
  }
}

/**
 * Duration formatter
 */
export class DurationFormatter {
  /**
   * Format milliseconds to human readable duration
   */
  static format(ms: number, options: { 
    includeMs?: boolean;
    maxUnits?: number;
    separator?: string;
  } = {}): string {
    const {
      includeMs = false,
      maxUnits = 3,
      separator = ' ',
    } = options;

    const units = [
      { label: 'd', value: 86400000 },
      { label: 'h', value: 3600000 },
      { label: 'm', value: 60000 },
      { label: 's', value: 1000 },
      { label: 'ms', value: 1 },
    ];

    const parts: string[] = [];
    let remaining = ms;

    for (const unit of units) {
      if (!includeMs && unit.label === 'ms') continue;
      
      if (remaining >= unit.value) {
        const count = Math.floor(remaining / unit.value);
        parts.push(`${count}${unit.label}`);
        remaining %= unit.value;
      }

      if (parts.length >= maxUnits) break;
    }

    return parts.join(separator) || '0s';
  }

  /**
   * Parse duration string to milliseconds
   */
  static parse(duration: string): number {
    const regex = /(\d+(?:\.\d+)?)\s*([dhms]|ms)/g;
    let match;
    let total = 0;

    while ((match = regex.exec(duration)) !== null) {
      const value = parseFloat(match[1]);
      const unit = match[2];

      switch (unit) {
        case 'd':
          total += value * 86400000;
          break;
        case 'h':
          total += value * 3600000;
          break;
        case 'm':
          total += value * 60000;
          break;
        case 's':
          total += value * 1000;
          break;
        case 'ms':
          total += value;
          break;
      }
    }

    return total;
  }
}

/**
 * Export all formatters
 */
export default {
  StringFormatter,
  TemplateFormatter,
  ByteFormatter,
  DurationFormatter,
};