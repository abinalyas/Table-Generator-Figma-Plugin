const faker = require('faker');

export function getFakerValue(type: string) {
    const t = type.toLowerCase();
    let result;
    switch (t) {
        case 'people name':
        case 'name':
            result = faker.name.findName();
            break;
        case 'first name':
            result = faker.name.firstName();
            break;
        case 'last name':
            result = faker.name.lastName();
            break;
        case 'brand name':
        case 'company':
            result = faker.company.companyName();
            break;
        case 'mobile number':
            result = faker.phone.phoneNumber();
            break;
        case 'date':
            result = faker.date.recent().toLocaleDateString();
            break;
        case 'random number':
        case 'number':
            result = faker.datatype.number({ min: 1, max: 1000 });
            break;
        case 'price':
            result = faker.commerce.price();
            break;
        case 'email':
            result = faker.internet.email();
            break;
        case 'product':
            result = faker.commerce.productName();
            break;
        case 'color':
            result = faker.commerce.color();
            break;
        default:
            result = faker.lorem.words(2);
    }
    return result;
}