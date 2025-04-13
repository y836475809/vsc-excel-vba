using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace VBALanguageServer {
	internal class DebounceDispatcher(int delay) {
		public bool IsCompleted;
		private int delay = delay;
		private CancellationTokenSource cancellationToken = null;

		public void Debounce(Action func) {
			this.IsCompleted = false;
			this.cancellationToken?.Cancel();
			this.cancellationToken = new CancellationTokenSource();

			Task.Delay(this.delay, this.cancellationToken.Token)
				.ContinueWith(task => {
					if (task.IsCompletedSuccessfully) {
						try {
							func();
							this.IsCompleted = true;
						} catch (Exception) {
							this.IsCompleted = true;
							throw;
						}
					}
				}, TaskScheduler.Default);
		}
	}
}
