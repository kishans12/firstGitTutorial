#include <winsock2.h>
#include <windows.h>
#include <stdexcept>
#include <iostream>
#include <memory>
#include <chrono>
#include <thread>
#include <cstdint>
#include <vector>
#include <evhttp.h>
#include<event2\thread.h>

#pragma comment(lib,"WS2_32")


class helloWorld
{
public:
	int increment;

	helloWorld()
	{
		this->increment = 0;
	}
	void setIncrement()
	{
		this->increment += 1;
	}

};



helloWorld lv_obj;

int main()
{
	
	char const SrvAddress[] = "127.0.0.1";
	std::uint16_t const SrvPort = 5555;
	int const SrvThreadCount = 2;

#ifdef _WIN32
	WORD wVersionRequested;
	WSADATA wsaData;

	wVersionRequested = MAKEWORD(2, 2);

	WSAStartup(wVersionRequested, &wsaData);
#endif
	evthread_use_windows_threads();

	try
	{
		void(*OnRequest)(evhttp_request *, void *) = [](evhttp_request *req, void *)
		{
			auto *OutBuf = evhttp_request_get_output_buffer(req);
			if (!OutBuf)
				return;

			//std::cout << "Thread Id: " << GetCurrentThreadId() << std::endl;
			std::this_thread::sleep_for(std::chrono::milliseconds(2000));			
			lv_obj.setIncrement();
			
			evbuffer_add_printf(OutBuf, "<html><body><center><h1>Hello World %d!</h1></center></body></html>", lv_obj.increment);
			//std::cout << "Request	"<<req<< std::endl;
			evhttp_send_reply(req, HTTP_OK, "", OutBuf);
			//std::cout << "Class Object" << lv_obj.increment<< std::endl;
			std::cout << "Response sent: " << lv_obj.increment << std::endl;
			std::cout << "Response send by Thread Id: " << GetCurrentThreadId() << std::endl;
		};
		std::exception_ptr InitExcept;
		bool volatile IsRun = true;
		evutil_socket_t Socket = -1;
		auto ThreadFunc = [&]()
		{
			std::cout<<"Thread id in ThreadFunc"<< std::this_thread::get_id()<<std::endl;
			try
			{
				std::unique_ptr<event_base, decltype(&event_base_free)> EventBase(event_base_new(), &event_base_free);
				if (!EventBase)
					throw std::runtime_error("Failed to create new base_event.");
				std::unique_ptr<evhttp, decltype(&evhttp_free)> EvHttp(evhttp_new(EventBase.get()), &evhttp_free);
				if (!EvHttp)
					throw std::runtime_error("Failed to create new evhttp.");
				evhttp_set_gencb(EvHttp.get(), OnRequest, nullptr);
				if (Socket == -1)
				{
					auto *BoundSock = evhttp_bind_socket_with_handle(EvHttp.get(), SrvAddress, SrvPort);
					if (!BoundSock)
						throw std::runtime_error("Failed to bind server socket.");
					if ((Socket = evhttp_bound_socket_get_fd(BoundSock)) == -1)
						throw std::runtime_error("Failed to get server socket for next instance.");
				}
				else
				{
					if (evhttp_accept_socket(EvHttp.get(), Socket) == -1)
						throw std::runtime_error("Failed to bind server socket for new instance.");
				}
				for (; IsRun; )
				{
					event_base_dispatch(EventBase.get());
					//event_base_loop(EventBase.get(),0);
					std::this_thread::sleep_for(std::chrono::milliseconds(10));
				}
			}
			catch (...)
			{
				InitExcept = std::current_exception();
			}
		};
		auto ThreadDeleter = [&](std::thread *t) { IsRun = false; t->join(); delete t; };
		typedef std::unique_ptr<std::thread, decltype(ThreadDeleter)> ThreadPtr;
		typedef std::vector<ThreadPtr> ThreadPool;
		ThreadPool Threads;
		for (int i = 0; i < SrvThreadCount; ++i)
		{
			ThreadPtr Thread(new std::thread(ThreadFunc), ThreadDeleter);
			std::this_thread::sleep_for(std::chrono::milliseconds(500));
			if (InitExcept != std::exception_ptr())
			{
				IsRun = false;
				std::rethrow_exception(InitExcept);
			}

			Threads.push_back(std::move(Thread));
		}
		std::cout << "Press Enter for quit." << std::endl;
		std::cin.get();
		IsRun = false;
	}
	catch (std::exception const &e)
	{
		std::cerr << "Error: " << e.what() << std::endl;
	}
	return 0;
}



