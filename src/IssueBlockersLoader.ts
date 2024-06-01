abstract class IssueBlockersLoader {

    loadBlockers(issueId: IssueId): Issue[] {
        return this.loadBlockersBulk([issueId])
    }

    loadBlockersBulk(issueIds: IssueId[]): Issue[] {
        return []
    }

}

abstract class IssueBlockersLoaderFactory {

    getIssueBlockerLoader(issueId: IssueId): IssueBlockersLoader | undefined {
        return undefined
    }

}
