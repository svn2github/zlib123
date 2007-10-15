#pragma once

class CSetLocale
{
public:
	CSetLocale(WORD wLandID) {
		m_old = GetThreadLocale();
		SetThreadLocale(MAKELCID(wLandID, 0));
	}

	~CSetLocale(void) {
		SetThreadLocale(m_old);
	}

private:
	DWORD m_old;
};
