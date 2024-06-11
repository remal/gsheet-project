abstract class IssueTracker {

    supportIssue(issueId: IssueId): boolean {
        throw Utils.throwNotImplemented(this.constructor.name, this.supportIssue.name)
    }

    canonizeIssueId(issueId: IssueId): IssueId {
        throw Utils.throwNotImplemented(this.constructor.name, this.canonizeIssueId.name)
    }

    getIssueLink(issueId: IssueId): string | null | undefined {
        throw Utils.throwNotImplemented(this.constructor.name, this.getIssueLink.name)
    }

    getIssuesLink(issueIds: IssueId[]): string | null | undefined {
        throw Utils.throwNotImplemented(this.constructor.name, this.getIssuesLink.name)
    }

    loadIssues(issueIds: IssueId[]): Record<IssueId, Issue | null | undefined> {
        throw Utils.throwNotImplemented(this.constructor.name, this.loadIssues.name)
    }

    loadChildren(issueIds: IssueId[]): Issue[] {
        throw Utils.throwNotImplemented(this.constructor.name, this.loadChildren.name)
    }

    loadBlockers(issueIds: IssueId[]): Issue[] {
        throw Utils.throwNotImplemented(this.constructor.name, this.loadBlockers.name)
    }

    search(query: string): Issue[] {
        throw Utils.throwNotImplemented(this.constructor.name, this.search.name)
    }

}

type IssueId = string

abstract class Issue {

    get id(): IssueId {
        throw Utils.throwNotImplemented(this.constructor.name, 'id')
    }

    get title(): string {
        throw Utils.throwNotImplemented(this.constructor.name, 'title')
    }

    get status(): string {
        throw Utils.throwNotImplemented(this.constructor.name, 'status')
    }

    get open(): boolean {
        throw Utils.throwNotImplemented(this.constructor.name, 'open')
    }

}

