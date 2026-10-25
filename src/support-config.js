// Paste an owner-confirmed PayPal link here, then rebuild the portable app.
// Empty or unsupported values hide the optional support link completely.
export const donationUrl = '';
export const label = '请我喝杯咖啡';

// Format validation only: this does not verify an account or make a request.
// PayPal.Me names: https://www.paypal.com/us/cshelp/article/help432
export function validateDonationUrl(value) {
  if (typeof value !== 'string') return '';
  const input = value.trim();
  if (!input || /[\u0000-\u0020\u007f\\]/.test(input)) return '';
  try {
    const url = new URL(input);
    if (url.protocol !== 'https:' || url.username || url.password || url.port || url.hash) return '';
    if (['paypal.me', 'www.paypal.me'].includes(url.hostname)) {
      if (!/^\/[a-z0-9]{1,20}\/?$/i.test(url.pathname)) return '';
      const seen = new Set();
      for (const [key, parameter] of url.searchParams) {
        if (seen.has(key)) return '';
        seen.add(key);
        if (key === 'country.x' && /^[a-z]{2}$/i.test(parameter)) continue;
        if (key === 'locale.x' && /^[a-z]{2,3}(?:[_-][a-z]{2})?$/i.test(parameter)) continue;
        return '';
      }
      url.hostname = 'paypal.me';
      url.search = '';
      return url.href;
    }
    if (url.hostname === 'www.paypal.com' && /^\/donate\/?$/.test(url.pathname)) {
      const parameters = [...url.searchParams];
      return parameters.length === 1 && parameters[0][0] === 'hosted_button_id'
        && /^[a-z0-9]{1,64}$/i.test(parameters[0][1]) ? url.href : '';
    }
  } catch { /* An invalid URL keeps the optional link hidden. */ }
  return '';
}
