// net/throttle.ts
export class Semaphore {
	private running = 0;
	private queue: Array<() => void> = [];
	constructor(private max = 6) {}

	acquire(): Promise<() => void> {
		return new Promise((resolve) => {
			const tryRun = () => {
				if (this.running < this.max) {
					this.running++;
					const release = () => {
						this.running = Math.max(0, this.running - 1);
						const job = this.queue.shift();
						if (job) job();
					};
					resolve(release);
				} else {
					this.queue.push(tryRun);
				}
			};
			tryRun();
		});
	}
}

export const globalSemaphore = new Semaphore(6);
