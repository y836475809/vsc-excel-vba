using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace VBALanguageServer {
	internal class DebounceDispatcher(int delay) {
		private int delay = delay;
		private CancellationTokenSource cancellationToken = null;
		private bool isCompleted = true;

		public bool IsCompleted {
			get { return this.isCompleted; }
		}

		public void Debounce(Action func) {
			this.isCompleted = false;
			this.cancellationToken?.Cancel();
			this.cancellationToken = new CancellationTokenSource();

			Task.Delay(this.delay, this.cancellationToken.Token)
				.ContinueWith(task => {
					if (task.IsCompletedSuccessfully) {
						func();
					}
					this.isCompleted = true;
				}, TaskScheduler.Default);
		}
	}
}
