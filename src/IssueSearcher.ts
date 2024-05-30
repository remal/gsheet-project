abstract class IssueSearcher {

    search(query: string): Issue[] {
        return []
    }

    canonizeQuery(query: string): string {
        return query
    }

    createWebUrl(query: string): string | null {
        return null
    }

}
