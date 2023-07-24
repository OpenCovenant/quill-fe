import { Component, OnInit } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { environment } from '../../environments/environment';
import {} from '../local-storage/local-storage.service';

@Component({
    selector: 'app-settings',
    templateUrl: './settings.component.html',
    styleUrls: ['./settings.component.css']
})
export class SettingsComponent implements OnInit {
    ALREADY_MADE_MARKING_TYPE_SELECTIONS =
        'penda-has-stored-marking-type-selections';

    baseURL!: string;
    markingTypesURL!: string;

    markingTypes: any[] = [];

    constructor(private http: HttpClient) {}

    ngOnInit(): void {
        this.initializeURLs();
        // TODO: consider making changes to `LocalStorageService`
        this.http.get(this.markingTypesURL).subscribe((data: any) => {
            if (
                !localStorage.getItem(this.ALREADY_MADE_MARKING_TYPE_SELECTIONS)
            ) {
                localStorage.setItem(
                    this.ALREADY_MADE_MARKING_TYPE_SELECTIONS,
                    'true'
                );
                this.markingTypes = Object.entries(
                    data['marking_types']
                ).filter((e: any) => e[1].enabled);
                this.markingTypes.forEach((mT) =>
                    localStorage.setItem(mT[0], mT[1].enabled)
                );
            } else {
                // TODO: lot of cases here in which more/less types come from the endpoint than are in local storage
                this.markingTypes = Object.entries(data['marking_types'])
                    .filter((e: any) => e[1].enabled)
                    .map((e: any) => {
                        e[1].enabled = localStorage.getItem(e[0]) === 'true';
                        return e;
                    });
            }
        });
    }

    onMarkingTypeSelection(markingTypeID: string, e: any): void {
        localStorage.setItem(markingTypeID, e);
    }

    private initializeURLs(): void {
        this.baseURL = environment.baseURL;
        this.markingTypesURL = `${this.baseURL}/api/getMarkingTypes`;
    }
}
