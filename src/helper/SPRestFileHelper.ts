// This has moved to /src/Common/js

//Author: Tim Wheeler (http://timwheeler.io)
//Purpose: Upload a file to a SharePoint library using REST and jQuery
interface IProgressCallback {
    (percentComplete: number): void;
}
export class SPRestFileUploader {
    public uploadFileAsArrayBuffer(file: File, webUrl: string, listname: string, filename: string, progressCallback: IProgressCallback = null): JQueryPromise<any> {
        var deferred = $.Deferred();
        this.getArrayBuffer(file)
            .fail((error) => {
                console.error("Get file buffer failed: " + error);
                deferred.reject(error);
            })
            .done((arrayBuffer: any) => {
                let options = {
                    url: `${webUrl}/_api/web/lists/getByTitle('${listname}')/RootFolder/Files/Add(url='${filename}', overwrite=true)`,
                    type: "POST",
                    data: arrayBuffer,
                    processData: false,
                    headers: {
                        "accept": "application/json;odata=verbose",
                        "X-RequestDigest": jQuery("#__REQUESTDIGEST").val(),
                    }
                } as any;
                $.ajax(options)
                    .fail((jqXHR, textStatus) => {
                        if (jqXHR && jqXHR.responseText && jqXHR.responseText.indexOf("The security validation for this page is invalid") > -1) {
                            console.error("SharePoint Page Validation has expired, please refresh.");
                            deferred.reject(jqXHR, "SharePoint Page Validation has expired, please save your work and refresh the page.");
                        }
                        else {
                            console.error("Upload file failed: " + textStatus);
                            deferred.reject(jqXHR, textStatus);
                        }
                    })
                    .done((data: any, textStatus: string, jqXHR: JQueryXHR) => {
                        if (data.d) {
                            deferred.resolve(data.d);
                        }
                        else {
                            deferred.resolve(data);
                        }
                    });
            });
        return deferred.promise();
    }
    public updateListItem(webUrl: string, serverRelativeUrl: string, properties: any, etag: string = "*") {
        if (etag) {
            etag = etag.replace(/{|}/g, "").toLowerCase();
        }

        if (typeof (properties.__metadata) == typeof (undefined)) {
            properties.__metadata = { 'type': 'SP.ListItem' };
        }
        var url = `${webUrl}/_api/Web/GetFileByServerRelativeUrl('${serverRelativeUrl}')/ListItemAllFields `;
        // This call does not return response content from the server.
        return $.ajax({
            url: url,
            type: "POST",
            data: JSON.stringify(properties),
            headers: {
                "X-RequestDigest": $("#__REQUESTDIGEST").val(),
                "content-type": "application/json;odata=verbose",
                "IF-MATCH": etag,
                "X-HTTP-Method": "MERGE"
            }
        } as any);
    }
    private getArrayBuffer(file: File) {
        var deferred = $.Deferred();
        var reader = new FileReader();
        reader.onloadend = (e: any) => {
            deferred.resolve(e.target.result);
        }
        reader.onerror = (e: any) => {
            deferred.reject(e.target.error);
        }
        reader.readAsArrayBuffer(file);
        return deferred.promise();
    }
}