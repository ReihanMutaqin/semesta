// Firebase Web SDK Configuration for Data Semesta
var firebaseConfig = window.FIREBASE_CONFIG || {
  apiKey: "AIzaSyA7FqfAqXaQl8c-Fbp2RACwa0XqfaGCGJM",
  authDomain: "data-semesta.firebaseapp.com",
  databaseURL: "https://data-semesta-default-rtdb.asia-southeast1.firebasedatabase.app",
  projectId: "data-semesta",
  storageBucket: "data-semesta.firebasestorage.app",
  messagingSenderId: "540118142385",
  appId: "1:540118142385:web:31849c95c9f2af3ac78190"
};

// Initialize Firebase
if (typeof firebase !== 'undefined') {
  if (!firebase.apps.length) {
    firebase.initializeApp(firebaseConfig);
  }
}
