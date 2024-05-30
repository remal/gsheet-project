abstract class IssueLoader {

    load(issueId: string): Issue | null {
        return null
    }

    canonizeId(issueId: string): string {
        return issueId
    }

    createWebUrl(issueId: string): string | null {
        return null
    }

}
