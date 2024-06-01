abstract class IssueChildrenLoader {

    loadChildren(issueId: IssueId): Issue[] {
        return this.loadChildrenBulk([issueId])
    }

    loadChildrenBulk(issueIds: IssueId[]): Issue[] {
        return []
    }

}

abstract class IssueChildrenLoaderFactory {

    getIssueChildrenLoader(issueId: IssueId): IssueChildrenLoader | undefined {
        return undefined
    }

}
