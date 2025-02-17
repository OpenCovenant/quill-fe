import { Component } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { DarkModeService } from '../services/dark-mode.service';
import { environment } from '../../environments/environment';
import { ActivatedRoute, Router } from '@angular/router';
import { CommonModule } from '@angular/common';

@Component({
    selector: 'app-terms-of-use',
    templateUrl: './terms-of-use.component.html',
    styleUrls: ['./terms-of-use.component.css'],
    imports: [CommonModule]
})
export class TermsOfUseComponent {
    baseURL!: string;
    getTermsOfUseURL!: string;

    constructor(
        public darkModeService: DarkModeService,
        private httpClient: HttpClient,
        private activatedRoute: ActivatedRoute,
        private router: Router
    ) {
        this.initializeURLs();
        this.httpClient.get(this.getTermsOfUseURL).subscribe((html: any) => {
            document.getElementById('terms-of-use')!.innerHTML =
                html['terms_of_use_html'];

            this.handleURLFragmenting();
        });
    }

    private handleURLFragmenting(): void {
        const fragment = this.router.url.split('#')[1];

        if (fragment) {
            this.activatedRoute.fragment.subscribe((f) => {
                document.querySelector('#' + f)?.scrollIntoView();
            });
        }

        document.querySelectorAll('.link-icon').forEach((lI) =>
            lI.addEventListener('click', (e) => {
                const currentFragment = (e.target as any).parentNode.id;
                window.location.href =
                    window.location.pathname +
                    window.location.search +
                    '#' +
                    currentFragment;
                document.querySelector('#' + currentFragment)?.scrollIntoView();
            })
        );
    }

    private initializeURLs(): void {
        this.baseURL = environment.baseURL;
        this.getTermsOfUseURL = this.baseURL + '/api/getTermsOfUse';
    }
}
