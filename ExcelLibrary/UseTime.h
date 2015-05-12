#pragma once

class UseTime
{
public:
	UseTime()
	{
		beginTime = ::GetTickCount();
	}

	~UseTime()
	{
		DWORD endTime = ::GetTickCount();
		std::cout << endTime - beginTime << std::endl;
	}

private:
	DWORD beginTime;

};
