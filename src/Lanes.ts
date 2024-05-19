type LaneAddCallback<T> = (laneIndex: number, lane: Lane<T>, amount: number, object: T) => void

class Lanes<T> {

    readonly lanes: Lane<T>[] = []

    constructor(lanesNumber: number) {
        for (let i = 0; i < lanesNumber; ++i) {
            this.lanes.push(new Lane<T>())
        }
    }

    add(amount: number, object: T, callback?: LaneAddCallback<T>) {
        if (!this.lanes.length) {
            return this
        }

        let laneIndex = 0
        let laneToAdd = this.lanes[laneIndex]
        let currentSum = laneToAdd.sum
        for (let i = 1; i < this.lanes.length; ++i) {
            const lane = this.lanes[i]
            const sum = lane.sum
            if (sum < currentSum) {
                laneIndex = i
                laneToAdd = lane
                currentSum = sum
            }
        }

        laneToAdd.add(amount, object)
        if (callback != null) {
            callback(laneIndex, laneToAdd, amount, object)
        }
        return this
    }

}
