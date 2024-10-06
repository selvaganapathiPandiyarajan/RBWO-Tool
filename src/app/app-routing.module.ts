import { NgModule } from '@angular/core';
import { RouterModule, Routes } from '@angular/router';
import { UploadsectionComponent } from './uploadsection/uploadsection.component';


const routes: Routes = [
  {path:'', redirectTo:'upload', pathMatch:'full'},
  {path:"upload",component:UploadsectionComponent},
];

@NgModule({
  imports: [RouterModule.forRoot(routes)],
  exports: [RouterModule]
})
export class AppRoutingModule { }
