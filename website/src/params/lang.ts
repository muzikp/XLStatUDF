export function match(param: string): param is 'cs' | 'en' {
  return param === 'cs' || param === 'en';
}

