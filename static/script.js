// script.js
var socket = io.connect('http://' + document.domain + ':' + location.port + '/test');

socket.on('connect', function() {
    console.log('Connected to server');
});

socket.on('conversion_started', function() {
    document.getElementById('conversion-status').innerHTML = 'Conversion in progress...';
});

socket.on('conversion_completed', function() {
    document.getElementById('conversion-status').innerHTML = 'Conversion completed!';
});