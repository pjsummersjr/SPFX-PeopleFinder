import Profile from "../models/Profile";
import { SPHttpClient, SPHttpClientConfiguration, SPHttpClientResponse } from "@microsoft/sp-http";

export interface IProfileService {
    getProfiles(): Promise<Profile[]>;
    searchProfiles(searchTerms: string): Promise<Profile[]>;
}

export class ProfileService implements IProfileService {

    private spClient: SPHttpClient;
    private webUrl: string;
    constructor(spclient: SPHttpClient, webUrl: string) {
        this.spClient = spclient;
        this.webUrl = webUrl;
    }

    public getProfiles(): Promise<Profile[]> {
        return this.spClient.get(this.webUrl + "/_api/search/query?QueryText='*'&SelectProperties='PreferredName,WorkEmail'&SourceId='b09a7990-05ea-4af9-81ef-edfab16c4e31'", SPHttpClient.configurations.v1,
        {headers: {"Accept": "application/json;odata=nometadata", "odata-version":""}})
        .then((response: SPHttpClientResponse) => {
            if(response.status !== 200) return Promise.reject(`An error occurred.\n${response.statusMessage}`);
            let someProfiles: Profile[] = [];
            return response.json().then((resultdata: any) => {
                if(resultdata.PrimaryQueryResult && resultdata.PrimaryQueryResult.RelevantResults 
                    && resultdata.PrimaryQueryResult.RelevantResults.Table
                ){
                    let resultRows = resultdata.PrimaryQueryResult.RelevantResults.Table.Rows;
                    resultRows.map((item) => {
                        let current: Profile = new Profile();
                        item.Cells.forEach((cell, index) => {
                            if(cell.Key === "PreferredName") current.displayName = cell.Value;
                            if(cell.Key === "WorkEmail") current.emailAddress = cell.Value;
                        })
                        someProfiles.push(current);
                    })                    
                }
                return someProfiles;                
            })
            
        },
        (error) => {
            return Promise.reject(`An error occurred:\n${error}`);
        }) 
    }

    public searchProfiles(searchTerms: string): Promise<Profile[]> {
        return Promise.reject({"error": "Method not yet implemented"});
    }
}

export class MockProfileService implements IProfileService {
    public getProfiles(): Promise<Profile[]> {
        return Promise.resolve([
           {
               "displayName": "Paul Summers",
               "emailAddress": "paul.j.summers@bc.edu",
               "phoneNumber": "555-555-1212",
               "title": "Chief Architect",
               "pictureUrl": "https://familyguyaddicts.files.wordpress.com/2014/04/peter-animation-033idlepic4x.png",
               "skills": [],
               "interests": [],
               "recentDocuments": []
           } 
        ]);
    }

    public searchProfiles(searchTerms: string): Promise<Profile[]> {
        return Promise.reject({"error": "Method not yet implemented"});
    }
}