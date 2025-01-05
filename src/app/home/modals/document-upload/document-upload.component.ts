import { Component, OnDestroy } from '@angular/core';
import { CommonModule } from '@angular/common';
import { HttpClient } from '@angular/common/http';
import { RouterLink } from '@angular/router';
import {
    BehaviorSubject,
    interval,
    map,
    Observable,
    startWith,
    Subject,
    Subscription,
    takeUntil
} from 'rxjs';
import * as XLSX from 'xlsx';
import { WorkBook, WorkSheet } from 'xlsx';

import { DarkModeService } from '../../../services/dark-mode.service';
import { environment } from '../../../../environments/environment';
import { MarkedPage } from '../../../models/marked-page';
import { DocumentMarking } from '../../../models/document-marking';

@Component({
    selector: 'app-document-upload',
    templateUrl: './document-upload.component.html',
    styleUrl: './document-upload.component.css',
    imports: [CommonModule, RouterLink]
})
export class DocumentUploadComponent implements OnDestroy {
    loading$: BehaviorSubject<boolean> = new BehaviorSubject<boolean>(false);
    markedPages: MarkedPage[] | undefined = undefined;
    totalMarkingsCount: number | undefined = undefined;
    totalPagesMarkedCount: number | undefined = undefined;
    processingTimeLeft: string | undefined = undefined;

    private baseURL!: string;
    private uploadDocumentURL!: string;
    private documentUploadSubscription$: Subscription | undefined = undefined;
    private countdownSubject: Subject<void> = new Subject();
    private countdown$: Observable<number> | undefined = undefined;

    constructor(
        public darkModeService: DarkModeService,
        private httpClient: HttpClient
    ) {
        this.initializeURLs();
    }

    ngOnDestroy(): void {
        this.documentUploadSubscription$?.unsubscribe();
    }

    /**
     * Uploads the selected document to be marked.
     * @param {Event} $event the event emitted when the file is selected
     */
    uploadDocument($event: Event): void {
        const fileList: FileList | null = ($event.target as HTMLInputElement)
            .files;
        if (fileList && fileList.length === 1) {
            const file: File = fileList[0];
            const fileSize = file.size / (1024 * 1024);
            if (fileSize > 30) {
                alert(
                    'Dokumenti është tepër i madh. Zgjidhni një dokument më të vogel se 30 MB.'
                );
                return;
            }
            this.clearModal();
            this.loading$.next(true);

            const sizeBasedTimerEstimation =
                (Math.round(fileSize) + 1) * 60 * 1000;
            this.startProcessingTimer(sizeBasedTimerEstimation);

            document.getElementById('uploaded-document-name')!.textContent =
                file.name;
            document.getElementById('uploaded-document-name')!.title =
                file.name;
            const formData: FormData = new FormData();
            formData.append('uploadFile', file, file.name);
            this.documentUploadSubscription$ = this.httpClient
                .post(this.uploadDocumentURL, formData)
                .subscribe({
                    next: (value) => {
                        const markedPages = value as MarkedPage[];
                        this.markedPages = markedPages;
                        this.totalMarkingsCount = markedPages
                            .map((mP) => mP.markings.length)
                            .reduce((a, b) => a + b, 0);
                        this.totalPagesMarkedCount = markedPages.length;
                        this.loading$.next(false);
                        this.stopProcessingTimer();
                    },
                    error: () => {
                        this.stopProcessingTimer();
                        this.loading$.next(false);
                        this.clearModal();
                        alert('Ngarkimi dështoi, provojeni përsëri.');
                    }
                });
        } else {
            alert('Ngarko vetëm një dokument!');
        }
    }

    clearModal(): void {
        this.markedPages = undefined;
        this.totalMarkingsCount = undefined;
        this.totalPagesMarkedCount = undefined;
        (document.getElementById('document-upload-input')! as any).value = null;
        document.getElementById('uploaded-document-name')!.textContent = '';
        document.getElementById('uploaded-document-name')!.title = '';
    }

    downloadReport(): void {
        const data: any[][] = [
            ['Numri i Faqes', 'Shenjimi', 'Lloji', 'Përshkrimi', 'Sugjerimet']
        ];

        this.markedPages?.forEach((markedPage: MarkedPage) => {
            markedPage.markings.forEach((pageMarking: DocumentMarking) => {
                const markingEntry = [
                    markedPage.pageNumber,
                    pageMarking.text,
                    pageMarking.subtype,
                    pageMarking.description,
                    pageMarking.suggestions
                ];
                data.push(markingEntry);
            });
        });

        const columns = Object.keys(data[0]);
        const columnWidths = columns.map((col) => ({
            wch: Math.max(col.length, this.getMaxCellLength(data, col))
        }));

        const markedPagesReport: WorkBook = XLSX.utils.book_new();
        const reportWorksheet: WorkSheet = XLSX.utils.aoa_to_sheet(data);
        reportWorksheet['!cols'] = columnWidths;

        XLSX.utils.book_append_sheet(
            markedPagesReport,
            reportWorksheet,
            'penda-raport'
        );

        const fileName = document.getElementById(
            'uploaded-document-name'
        )!.textContent;
        XLSX.writeFile(
            markedPagesReport,
            `penda-${fileName?.substring(0, fileName?.length - 4)}.xlsx`
        );
    }

    private startProcessingTimer(totalMinutes: number = 3 * 60 * 1000): void {
        this.countdown$ = interval(60 * 1000).pipe(
            startWith(0),
            map((_, index) => totalMinutes - index * 60 * 1000),
            takeUntil(this.countdownSubject)
        );

        this.countdown$.subscribe((count: number) => {
            const countInMinutes = count / 60000;
            this.processingTimeLeft =
                countInMinutes === 1 ? '1 minutë' : `${countInMinutes} minuta`;
            if (count === 0) {
                this.stopProcessingTimer();
            }
        });
    }

    private stopProcessingTimer(): void {
        this.countdownSubject.next();
    }

    private getMaxCellLength(arr: any[], property: string): number {
        return Math.max(...arr.map((item) => item[property].toString().length));
    }

    private initializeURLs(): void {
        this.baseURL = environment.baseURL;
        this.uploadDocumentURL = this.baseURL + '/api/uploadDocument';
    }
}
