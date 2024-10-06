import { Pipe, PipeTransform } from '@angular/core';

@Pipe({
  name: 'custom'
})
export class CustomPipe implements PipeTransform {

  transform(value: Date, ...args: any[]): string {
    let options:any = {month: 'short', year: 'numeric'};
    return new Intl.DateTimeFormat('en-US', options).format(value);

  }
}
