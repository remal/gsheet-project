type IssueKey = string
type IssueId = string
type IssueSearchQuery = string

abstract class IssueTracker {

    supportsIssueKey(issueKey: IssueKey): boolean {
        return this.extractIssueId(issueKey) != null
            || this.extractSearchQuery(issueKey) != null
    }

    canonizeIssueKey(issueKey: IssueKey) {
        {
            const issueId = this.extractIssueId(issueKey)
            if (issueId?.length) {
                const canonizedKey = this.issueIdToIssueKey(issueId)
                if (canonizedKey?.length) {
                    return canonizedKey
                }
            }
        }

        {
            const searchQuery = this.extractSearchQuery(issueKey)
            if (searchQuery?.length) {
                const canonizedKey = this.searchQueryToIssueKey(searchQuery)
                if (canonizedKey?.length) {
                    return canonizedKey
                }
            }
        }

        return issueKey
    }


    extractIssueId(issueKey: IssueKey): IssueId | null | undefined {
        throw Utils.throwNotImplemented(this.constructor.name, this.extractIssueId.name)
    }

    issueIdToIssueKey(issueId: IssueId): IssueKey | null | undefined {
        throw Utils.throwNotImplemented(this.constructor.name, this.issueIdToIssueKey.name)
    }

    getUrlForIssueId(issueId: IssueId): string | null | undefined {
        return this.getUrlForIssueIds([issueId])
    }

    getUrlForIssueIds(issueIds: IssueId[]): string | null | undefined {
        if (!issueIds?.length) {
            return undefined
        }

        throw Utils.throwNotImplemented(this.constructor.name, this.getUrlForIssueIds.name)
    }

    loadIssues(issueIds: IssueId[]): Issue[] {
        if (!issueIds?.length) {
            return []
        }

        throw Utils.throwNotImplemented(this.constructor.name, this.loadIssues.name)
    }

    loadChildren(issueIds: IssueId[]): Issue[] {
        if (!issueIds?.length) {
            return []
        }

        throw Utils.throwNotImplemented(this.constructor.name, this.loadChildren.name)
    }

    loadBlockers(issueIds: IssueId[]): Issue[] {
        if (!issueIds?.length) {
            return []
        }

        throw Utils.throwNotImplemented(this.constructor.name, this.loadBlockers.name)
    }


    extractSearchQuery(issueKey: IssueKey): IssueSearchQuery | null | undefined {
        throw Utils.throwNotImplemented(this.constructor.name, this.extractSearchQuery.name)
    }

    searchQueryToIssueKey(query: IssueSearchQuery): IssueKey | null | undefined {
        throw Utils.throwNotImplemented(this.constructor.name, this.issueIdToIssueKey.name)
    }

    getUrlForSearchQuery(query: IssueSearchQuery): string | null | undefined {
        throw Utils.throwNotImplemented(this.constructor.name, this.getUrlForSearchQuery.name)
    }

    loadIssueKeySearchTitle(issueKey: IssueKey): string | null | undefined {
        return this.extractSearchQuery(issueKey)
    }

    search(query: IssueSearchQuery): Issue[] {
        if (!query?.length) {
            return []
        }

        throw Utils.throwNotImplemented(this.constructor.name, this.search.name)
    }

}

abstract class Issue {

    readonly issueTracker: IssueTracker

    protected constructor(issueTracker: IssueTracker) {
        this.issueTracker = issueTracker
    }

    get id(): IssueId {
        throw Utils.throwNotImplemented(this.constructor.name, 'id')
    }

    get title(): string {
        throw Utils.throwNotImplemented(this.constructor.name, 'title')
    }

    get type(): string {
        throw Utils.throwNotImplemented(this.constructor.name, 'type')
    }

    get status(): string {
        throw Utils.throwNotImplemented(this.constructor.name, 'status')
    }

    get open(): boolean {
        throw Utils.throwNotImplemented(this.constructor.name, 'open')
    }

    get assignee(): string {
        throw Utils.throwNotImplemented(this.constructor.name, 'assignee')
    }

}

