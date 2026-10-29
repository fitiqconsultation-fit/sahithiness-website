<?php
// Razorpay webhook relay for sahithiness.in
// Checks the Razorpay signature, then forwards the payment to the Sahithiness Clients Google Sheet.
// Secrets live OUTSIDE public_html in sahithiness-config.php (never in GitHub):
//   <?php return ['webhook_secret' => '...', 'sheet_webhook_url' => 'https://script.google.com/macros/s/.../exec?token=...'];

header('Content-Type: application/json');
if ($_SERVER['REQUEST_METHOD'] !== 'POST') { http_response_code(405); echo '{"ok":false}'; exit; }

$cfgFile = dirname(__DIR__) . '/sahithiness-config.php';
if (!is_file($cfgFile)) { http_response_code(500); echo '{"ok":false,"error":"config missing"}'; exit; }
$cfg = require $cfgFile;

$body = file_get_contents('php://input');
$sig = isset($_SERVER['HTTP_X_RAZORPAY_SIGNATURE']) ? $_SERVER['HTTP_X_RAZORPAY_SIGNATURE'] : '';
$expected = hash_hmac('sha256', $body, $cfg['webhook_secret']);
if (!$sig || !hash_equals($expected, $sig)) { http_response_code(400); echo '{"ok":false,"error":"bad signature"}'; exit; }

$ch = curl_init($cfg['sheet_webhook_url']);
curl_setopt_array($ch, [
  CURLOPT_POST => true,
  CURLOPT_POSTFIELDS => $body,
  CURLOPT_HTTPHEADER => ['Content-Type: application/json'],
  CURLOPT_RETURNTRANSFER => true,
  CURLOPT_FOLLOWLOCATION => true,   // Google answers with a redirect; the script has already run by then
  CURLOPT_TIMEOUT => 25,
]);
$res = curl_exec($ch);
$code = curl_getinfo($ch, CURLINFO_HTTP_CODE);
curl_close($ch);

$data = json_decode((string)$res, true);
if ($code === 200 && is_array($data) && !empty($data['ok'])) { echo '{"ok":true}'; exit; }
http_response_code(502);   // Razorpay will retry later
echo '{"ok":false}';
