// التحقق من توفر المكتبات المطلوبة
function checkRequiredLibraries() {
    console.log('جاري التحقق من المكتبات...');
    console.log('XLSX:', typeof window.XLSX);
    
    const missingLibraries = [];

    if (typeof window.XLSX === 'undefined') {
        console.warn('⚠️ XLSX لم يتم تحميلها بعد');
        missingLibraries.push('XLSX (Excel)');
    } else {
        console.log('✅ XLSX محملة بنجاح');
    }

    if (missingLibraries.length > 0) {
        console.warn('المكتبات المفقودة:', missingLibraries);
        // محاولة إعادة التحقق بعد وقت
        setTimeout(checkRequiredLibraries, 1000);
        return false;
    }

    console.log('✅ جميع المكتبات المطلوبة متوفرة');
    return true;
}

// انتظار تحميل جميع المكتبات
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => {
        console.log('DOM تم تحميله، جاري التحقق من المكتبات...');
        setTimeout(checkRequiredLibraries, 1000);
    });
} else {
    console.log('DOM جاهز، جاري التحقق من المكتبات...');
    setTimeout(checkRequiredLibraries, 1000);
}
