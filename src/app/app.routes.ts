import { Routes } from '@angular/router';
import { pagesRoutes } from './core/pages/pages.routes';
import { PagesComponent } from './core/pages/pages.component';
// import { NotFoundComponent } from './not-found.component';

export const routes: Routes = [
    // Si la ruta es root '', redirigimos a las rutas de pages
    {
        path: '',
        component: PagesComponent,   // Layout principal con header/footer
        children: pagesRoutes,
    },

    // Wildcard: cualquier ruta que no exista
    //   { path: '**', component: NotFoundComponent },
];