abstract class IssueLoader {

    load(issueId: IssueId): Issue | null {
        return null
    }

    canonizeId(issueId: IssueId): IssueId {
        return issueId
    }

    createWebUrl(issueId: IssueId): string | null {
        return null
    }

}

abstract class IssueLoaderFactory {

    getIssueLoader(issueId: IssueId): IssueLoader | undefined {
        return undefined
    }

}
