// ضع هنا بيانات مشروع Firebase الخاصة بك (Project settings → Your apps → Web app → Config)
window.FIREBASE_CONFIG = {
    apiKey: "AIzaSyAjhkZ3q6d_L5RkBk_P2X1_PKbgdFe9LCo",
    authDomain: "cars-775f3.firebaseapp.com",
    projectId: "cars-775f3",
    storageBucket: "cars-775f3.firebasestorage.app",
    messagingSenderId: "337903607077",
    appId: "1:337903607077:web:c8f4e729b2dfd17b8ececc",
    measurementId: "G-07RNTMY9SS"
};

// false = كل مستخدم له بياناته الخاصة (كل يوزر وداتاه)
// true  = كل المستخدمين يشوفوا ويعدلوا نفس الداتا (داتا مشتركة للمصنع)
window.SHARED_DATA = false;

// يتم إضافته تلقائيًا لاسم المستخدم لو مكتبتش إيميل (لازم يطابق الإيميلات اللي عملتها في Firebase)
window.LOGIN_EMAIL_DOMAIN = "bahnasawy.app";
