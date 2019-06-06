#include <winsock2.h>
#include <windows.h>
#include<stdio.h>
#include <csignal>
#include <condition_variable>
#include "HttpServer.h"



#pragma comment(lib,"WS2_32")

std::mutex m;
std::condition_variable cv;

void signalHandler(int signum) 
{
	std::cout << "Sighandler" << std::endl;
	std::cout << "Signal (" << signum << ") received.\n";
	cv.notify_all();


}

void onHello(CHttpRequest *req, void *arg) 
{
	std::this_thread::sleep_for(std::chrono::milliseconds(5000));
	//Sleep(500);
	//for (int i = 0; i < 500000; i++) {
	//	for (int j = 0; j < 5000; j++) {
	//		//for (int k = 0; k < 50000000; k++) {
	//		//}
	//	}
	//}
	//std::cout<<"Check thread is enabled or not------>	"<<(int)EVTHREAD_GET_ID()<<std::endl;
	//std::cout << "Thread id in ThreadFunc" << std::this_thread::get_id() << std::endl;//Kishan
	char data[100] = {0};
	sprintf(data, "<html><body><center><h1>Hello World %d</h1></center></body></html>", std::this_thread::get_id());
	//req->AddBufferOut("hello World", 16);
	req->AddBufferOut(data, strlen(data));
	
	//std::cout << "---------Send Reply---------" << std::endl;
	req->SendReply(HTTP_OK);//evhttp_send_reply(req, HTTP_OK, "", OutBuf);

}

int main(int argc, char** argv)
{
	//event_enable_debug_logging(EVENT_DBG_ALL);
	WORD wVersionRequested;
	WSADATA wsaData;

	wVersionRequested = MAKEWORD(2, 2);
	WSAStartup(wVersionRequested, &wsaData);

	signal(SIGINT, signalHandler);
	signal(SIGTERM, signalHandler);
	std::cout << "Before Server object" << std::endl;
	CHttpServer server("127.0.0.1", 5555, 5);
	std::cout << "After Server object" << std::endl;
	server.SetCallback("/hello", onHello, NULL);
	
	std::cout << "After Server SetCallback" << std::endl;
	server.Start();
	
	//std::cout << "After Start" << std::endl;
	
	std::unique_lock<std::mutex> lk(m);
	cv.wait(lk);
	//server.Stop();
	return 0;
}