class IssueTrackerExample extends IssueTracker {

    issueIdToIssueKey(issueId: IssueId): IssueKey | null | undefined {
        return `example/${issueId}`
    }

    extractIssueId(issueKey: IssueKey): IssueId | null | undefined {
        return Utils.extractRegex(issueKey, /^example\/([\d.-]+)$/, 1)
    }

    getUrlForIssueId(issueId: IssueId): string | null | undefined {
        return `https://example.com/issues/${encodeURIComponent(issueId)}`
    }

    getUrlForIssueIds(issueIds: IssueId[]): string | null | undefined {
        if (!issueIds?.length) {
            return null
        }

        return `https://example.com/search?q=id:(${encodeURIComponent(issueIds.join('|'))})`
    }

    loadIssues(issueIds: IssueId[]): Issue[] {
        if (!issueIds?.length) {
            return []
        }

        return issueIds.map(id => new IssueExample(this, id))
    }

    loadChildren(issueIds: IssueId[]): Issue[] {
        if (!issueIds?.length) {
            return []
        }

        return issueIds.flatMap(id => {
            let hash = parseInt(id)
            if (isNaN(hash)) {
                hash = Math.abs(Utils.hashCode(id))
            }
            return Array.from(Utils.range(0, hash % 3)).map(index =>
                new IssueExample(this, `${id}-${index + 1}`),
            )
        })
    }

    loadBlockers(issueIds: IssueId[]): Issue[] {
        if (!issueIds?.length) {
            return []
        }

        return issueIds.flatMap(id => {
            let hash = parseInt(id)
            if (isNaN(hash)) {
                hash = Math.abs(Utils.hashCode(id))
            }
            return Array.from(Utils.range(0, hash % 2)).map(index =>
                new IssueExample(this, `${id}-blocker-${index + 1}`),
            )
        })
    }


    extractSearchQuery(issueKey: IssueKey): IssueSearchQuery | null | undefined {
        return Utils.extractRegex(issueKey, /^example\/search\/(.+)$/, 1)
    }

    searchQueryToIssueKey(query: IssueSearchQuery): IssueKey | null | undefined {
        return `example/search/${query}`
    }

    getUrlForSearchQuery(query: IssueSearchQuery): string | null | undefined {
        return `https://example.com/search?q=${encodeURIComponent(query)}`
    }

    search(query: IssueSearchQuery): Issue[] {
        if (!query?.length) {
            return []
        }

        const hash = Math.abs(Utils.hashCode(query))
        return Array.from(Utils.range(0, hash % 3)).map(index =>
            new IssueExample(this, `search-${hash}-${index + 1}`),
        )
    }

}

GSheetProjectSettings.issueTrackers.push(new IssueTrackerExample())


class IssueExample extends Issue {

    private readonly _id: IssueKey

    constructor(issueTracker: IssueTracker, id: IssueKey) {
        super(issueTracker)
        this._id = id
    }

    get id(): IssueKey {
        return this._id
    }

    get title(): string {
        return `Issue '${this.id}'`
    }

    get type(): string {
        return 'task'
    }

    get status(): IssueExampleStatus {
        let hash = parseInt(this.id)
        if (isNaN(hash)) {
            hash = Math.abs(Utils.hashCode(this.id))
        }
        return hash % 3 !== 0
            ? 'open'
            : 'closed';
    }

    get open(): boolean {
        return this.status === 'open'
    }

    get assignee(): string {
        let hash = parseInt(this.id)
        if (isNaN(hash)) {
            hash = Math.abs(Utils.hashCode(this.id))
        }
        return hash.toString()
    }

}

type IssueExampleStatus = 'open' | 'closed'
