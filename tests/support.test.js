import test from 'node:test';
import assert from 'node:assert/strict';
import { donationUrl, label, validateDonationUrl } from '../src/support-config.js';

test('support configuration is either empty or valid, and empty values stay hidden', () => {
  assert.ok(donationUrl === '' || validateDonationUrl(donationUrl), 'Configured support URL must use a supported PayPal format');
  assert.equal(typeof label, 'string');
  assert.ok(label.trim().length > 0);
  for (const value of ['', '   ', null, undefined, {}, 123]) assert.equal(validateDonationUrl(value), '');
});

test('support accepts only supported recipient-specific HTTPS link formats', () => {
  // Synthetic syntax fixtures only; these links are never requested.
  for (const value of [
    'https://paypal.me/ExampleCreator123',
    'https://paypal.me/ExampleCreator123/',
    'https://www.paypal.com/donate?hosted_button_id=EXAMPLE1234567',
    'https://www.paypal.com/donate/?hosted_button_id=EXAMPLE1234567',
  ]) assert.equal(validateDonationUrl(value), value);
  assert.equal(validateDonationUrl('  https://PAYPAL.ME/ExampleCreator123  '), 'https://paypal.me/ExampleCreator123');
  assert.equal(validateDonationUrl('https://www.paypal.me/ExampleCreator123'), 'https://paypal.me/ExampleCreator123');
  assert.equal(validateDonationUrl('https://www.paypal.me/ExampleCreator123?country.x=GB&locale.x=en_GB'), 'https://paypal.me/ExampleCreator123');
  assert.equal(validateDonationUrl('https://paypal.me/ExampleCreator123?locale.x=zh-CN'), 'https://paypal.me/ExampleCreator123');
});

test('support rejects unsafe schemes, lookalike hosts, credentials and unusual ports', () => {
  for (const value of [
    'javascript:alert(1)', 'data:text/html,synthetic', 'file:///tmp/donate',
    'http://paypal.me/ExampleCreator123', '//paypal.me/ExampleCreator123',
    'https://paypal.me.evil.invalid/ExampleCreator123',
    'https://evilpaypal.me/ExampleCreator123',
    'https://www.paypal.com.evil.invalid/donate?hosted_button_id=EXAMPLE1234567',
    'https://paypal.me@evil.invalid/ExampleCreator123',
    'https://evil@paypal.me/ExampleCreator123',
    'https://paypal.me:444/ExampleCreator123',
    'https://paypal.me/Example\nCreator123',
    'https://paypal.me\\ExampleCreator123',
  ]) assert.equal(validateDonationUrl(value), '', value);
});

test('support rejects generic pages, amounts, redirects and ambiguous recipient parameters', () => {
  for (const value of [
    'https://paypal.me/', 'https://paypal.me/ThisNameIsLongerThanTwentyCharacters',
    'https://paypal.me/Example-Creator', 'https://paypal.me/Example_Creator',
    'https://paypal.me/%45xampleCreator123', 'https://paypal.me/ExampleCreator123/25',
    'https://paypal.me/ExampleCreator123?redirect=https://evil.invalid',
    'https://paypal.me/ExampleCreator123?country.x=GB&country.x=CN',
    'https://paypal.me/ExampleCreator123?country.x=https://evil.invalid',
    'https://paypal.me/ExampleCreator123?locale.x=en_GB%26redirect%3Devil',
    'https://paypal.me/ExampleCreator123?country.x=GB&amount=25',
    'https://paypal.me/ExampleCreator123#fragment',
    'https://www.paypal.com/', 'https://www.paypal.com/donate',
    'https://www.paypal.com/donate?hosted_button_id=',
    'https://www.paypal.com/donate?hosted_button_id=ID1&hosted_button_id=ID2',
    'https://www.paypal.com/donate?hosted_button_id=EXAMPLE1234567&return=https://evil.invalid',
    'https://www.paypal.com/donate?business=someone%40example.invalid',
  ]) assert.equal(validateDonationUrl(value), '', value);
});
