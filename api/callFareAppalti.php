<?php

$curl = curl_init();
// echo 'Curl error: ' . curl_error($curl);

curl_setopt_array($curl, array(
  CURLOPT_URL => 'https://api.appaltando.it/API/v1/login.php?username=nextbroker&password=0cc80b3e231b5281dfef46b82eafbb30461cce34cc7e4c35779b1c434db2516c5de0081b77efef7691e9f2f0e952a1e4aed7e03c3b4fe066cebf3235c1049a86',
  CURLOPT_RETURNTRANSFER => true,
  CURLOPT_ENCODING => '',
  CURLOPT_MAXREDIRS => 10,
  CURLOPT_TIMEOUT => 0,
  CURLOPT_FOLLOWLOCATION => true,
  CURLOPT_HTTP_VERSION => CURL_HTTP_VERSION_1_1,
  CURLOPT_CUSTOMREQUEST => 'POST',
  CURLOPT_SSL_VERIFYHOST => 0,
  CURLOPT_SSL_VERIFYPEER => 0  
));

$response = curl_exec($curl);
// echo 'Curl error: ' . curl_error($curl);
curl_close($curl);
// echo $response;
$obj = json_decode($response);
$token = $obj->jwt;
echo 'token = ' . $token;


$curl = curl_init();

$headers = [
    'Content-Type: application/x-www-form-urlencoded; charset=utf-8',
    'Authorization: Token ' . $token
];


curl_setopt_array($curl, array(
  CURLOPT_URL => 'https://api.appaltando.it/API/v1/bandi.php?ente=ferrovia',
  CURLOPT_RETURNTRANSFER => true,
  CURLOPT_ENCODING => '',
  CURLOPT_MAXREDIRS => 10,
  CURLOPT_TIMEOUT => 0,
  CURLOPT_FOLLOWLOCATION => true,
  CURLOPT_HTTP_VERSION => CURL_HTTP_VERSION_1_1,
  CURLOPT_CUSTOMREQUEST => 'POST',
  CURLOPT_SSL_VERIFYHOST => 0,
  CURLOPT_SSL_VERIFYPEER => 0,
  CURLOPT_HTTPHEADER => $headers  
));
$response = curl_exec($curl);
echo 'Curl error: ' . curl_error($curl);
curl_close($curl);
echo $response;

?>