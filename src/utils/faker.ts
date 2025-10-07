// Faker functionality temporarily disabled for performance
// const faker = require('faker');

export function getFakerValue(type: string) {
    // Faker disabled - return simple placeholder values
    const t = type.toLowerCase();
    let result;
    switch (t) {
        case 'people name':
        case 'name':
            result = 'John Doe';
            break;
        case 'first name':
            result = 'John';
            break;
        case 'last name':
            result = 'Doe';
            break;
        case 'brand name':
        case 'company':
            result = 'Company Inc';
            break;
        case 'mobile number':
            result = '+1 (555) 123-4567';
            break;
        case 'date':
            result = new Date().toLocaleDateString();
            break;
        case 'random number':
        case 'number':
            result = Math.floor(Math.random() * 1000) + 1;
            break;
        case 'price':
            result = '$' + (Math.random() * 100).toFixed(2);
            break;
        case 'email':
            result = 'user@example.com';
            break;
        case 'product':
            result = 'Sample Product';
            break;
        case 'color':
            result = 'Blue';
            break;
        default:
            result = 'Sample Text';
    }
    return result;
}