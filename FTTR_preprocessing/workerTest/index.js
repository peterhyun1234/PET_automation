// index.js 파일
const { Worker } = require('worker_threads');
let startTime = process.uptime();           // 프로세스 시작 시간
let jobSize = 10000000000000;
let myWorker1, myWorker2;



// with worker thread


myWorker1 = new Worker(__dirname + '/w_threadTest.js');        // 스레드를 생성해 파일 절대경로를 통해 가리킨 js파일을 작업
myWorker2 = new Worker(__dirname + '/w_threadTest.js');

doSomething();  // event loop가 처리해야 할 CPU 하드한 작업

let endTime = process.uptime();
console.log("main thread time: " + (endTime - startTime));  // 스레드 생성 시간 + doSomething 처리하는 데 걸린 시간.

function doSomething() {
	let data;
	for (let i = 0; i < jobSize; i++) {      // CPU Hard
		data += i;
	}
}



// without worker thread

// 결론:
// CPU hard한 작업이 아니면 스레드 생성때문에 오히려 처리 시간이 늘어날 수 있다.
// 스레드를 생성하는 작업 자체가 부하가 크다고 생각함




// let jobSize = 30000000000000;                       // 이전 작업량의 3배

// // myWorker1 = new Worker(__dirname + '/worker.js');
// // myWorker2 = new Worker(__dirname + '/worker.js');

// doSomething();  // event loop가 처리해야 할 CPU 하드한 작업

// let endTime = process.uptime();
// console.log("main thread time: " + (endTime - startTime));  // 스레드 생성 시간 + doSomething 처리하는 데 걸린 시간.

// function doSomething() {
// 	let data;
// 	for (let i = 0; i < jobSize; i++) {      // CPU Hard
// 		data += i;
// 	}
// }