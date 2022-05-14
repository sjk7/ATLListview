// This is an independent project of an individual developer. Dear PVS-Studio,
// please check it.

// PVS-Studio Static Code Analyzer for C, C++, C#, and Java:
// http://www.viva64.com
#pragma once

#include <mmsystem.h>
struct SwInfo {
    DWORD m_start;
    DWORD m_stop;
    DWORD m_elapsed;
    bool armed;
};
namespace my {
class Stopwatch {
    public:
    Stopwatch(const CString& s, bool autostart = true) : m_string(s) {

        memset(&m_info, 0, sizeof(m_info));
        m_info.armed = true;
        if (autostart) {
            m_info.m_start = ::timeGetTime();
        }
    }

    ~Stopwatch() {
        if (m_info.armed && !m_string.IsEmpty()) {
            long el = (long)Elapsed();
            ATLTRACE(_T("******** %s took %ld ms. *******\n\n"),
                (LPCTSTR)m_string, el);
        }
    }

    __inline DWORD Elapsed() {
        if (m_info.m_start)
            return ::timeGetTime() - m_info.m_start;
        else
            return 0;
    }

    __inline void Start() { m_info.m_start = ::timeGetTime(); }
    __inline void Stop() { m_info.m_stop = ::timeGetTime(); }
    __inline bool isArmed() const { return m_info.armed; }
    __inline void Armed(bool b) { m_info.armed = b; }

    protected:
    SwInfo m_info;
    CString m_string;
};
} // namespace my
