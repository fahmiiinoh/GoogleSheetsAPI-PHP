<?php

require __DIR__ . '/vendor/autoload.php';

use Google\Client;

session_start();

getClientAuth();
/**
 * Returns an authorized API client.
 * @return Client the authorized client object
 */
function getClientAuth()
{
    $client = new Client();
    $client->setScopes('https://www.googleapis.com/auth/spreadsheets');
    $client->setAuthConfig('credentials.json');
    $client->setAccessType('offline');
    $client->setPrompt('select_account consent');

    // Load previously authorized token from cookie, if it exists.
    // The cookie named google_auth_token stores the user's access and refresh tokens, and is
    // created automatically when the authorization flow completes for the first
    // time.

        if (isset($_COOKIE['google_auth_token'])) {
        $accessToken = json_decode($_COOKIE['google_auth_token'], true);
        $client->setAccessToken($accessToken);
    
        } elseif (empty($_GET["code"])){

            $authUrl = $client->createAuthUrl();
            header('Location:'.$authUrl);

        } else {
            $authCode = $_GET["code"];
            $accessToken = $client->fetchAccessTokenWithAuthCode($authCode);
            $accessToken = $client->getAccessToken();
            
            $cookie_name = 'google_auth_token';
            $cookie_value = $accessToken;
            setcookie($cookie_name,json_encode($cookie_value), time() + 7200, "/");
            $accessToken = $_COOKIE['google_auth_token'];
             }   
            echo "<script>window.close();</script>";
            return $client;
         }