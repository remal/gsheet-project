class IssueTrackerExample extends IssueTracker {

    supportIssue(issueId: IssueId): boolean {
        return issueId.match(/^example\/\S+$/) != null
    }

    canonizeIssueId(issueId: IssueId): IssueId {
        return issueId
    }

    getIssueLink(issueId: IssueId): string | null | undefined {
        const searchQuery = issueId.match(/^example\/search\/(.*)$/)
        if (searchQuery != null) {
            return `https://example.com/search/?query=${encodeURIComponent(searchQuery[1])}`
        }

        return `https://example.com/issues/${encodeURIComponent(issueId)}`
    }

    getIssuesLink(issueIds: IssueId[]): string | null | undefined {
        return `https://example.com/search?query=id:(${encodeURIComponent(issueIds.join('|'))})`
    }

    loadIssues(issueIds: IssueId[]): Record<IssueId, Issue | null | undefined> {
        const result: Record<IssueId, Issue | null | undefined> = {}
        issueIds
            .filter(id => id?.length)
            .filter(Utils.distinct())
            .forEach(id => result[id] = new IssueExample(id))
        return result
    }

    loadChildren(issueIds: IssueId[]): Issue[] {
        return issueIds
            .filter(id => id?.length)
            .filter(Utils.distinct())
            .flatMap(id => {
                const children: Issue[] = []
                const childrenCount = Math.abs(Utils.hashCode(id)) % 3
                for (let child = 1; child <= childrenCount; ++child) {
                    children.push(new IssueExample(`${id}-${child}`))
                }
                return children
            })
    }

    loadBlockers(issueIds: IssueId[]): Issue[] {
        return []
    }

    search(query: string): Issue[] {
        const children: Issue[] = []
        const queryHash = Math.abs(Utils.hashCode(query))
        const childrenCount = queryHash % 3
        for (let child = 1; child <= childrenCount; ++child) {
            children.push(new IssueExample(`${queryHash}-${child}`))
        }
        return children
    }

}

class IssueExample extends Issue {

    private readonly _id: IssueId

    constructor(id: IssueId) {
        super()
        this._id = id
    }

    get id(): IssueId {
        return this._id
    }

    get title(): string {
        return `Issue '${this.id}'`
    }

    get status(): IssueExampleStatus {
        return Utils.hashCode(this.id) % 2 === 0
            ? 'open'
            : 'closed';
    }

    get open(): boolean {
        return this.status === 'open'
    }

}

type IssueExampleStatus = 'open' | 'closed'

GSheetProjectSettings.issueTrackers.push(new IssueTrackerExample())
