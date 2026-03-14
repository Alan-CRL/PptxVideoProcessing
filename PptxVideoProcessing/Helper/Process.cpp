module;

#define NOMINMAX
#include <Windows.h>

module PptxVideoProcessing.Helper.Process;

import std.compat;

import PptxVideoProcessing.Helper.Utf;

namespace
{
    class UniqueHandle
    {
    public:
        UniqueHandle() = default;
        explicit UniqueHandle(HANDLE handle) noexcept : handle_(handle)
        {
        }

        UniqueHandle(const UniqueHandle&) = delete;
        UniqueHandle& operator=(const UniqueHandle&) = delete;

        UniqueHandle(UniqueHandle&& other) noexcept : handle_(std::exchange(other.handle_, nullptr))
        {
        }

        UniqueHandle& operator=(UniqueHandle&& other) noexcept
        {
            if (this != &other)
            {
                reset();
                handle_ = std::exchange(other.handle_, nullptr);
            }

            return *this;
        }

        ~UniqueHandle()
        {
            reset();
        }

        [[nodiscard]] HANDLE get() const noexcept
        {
            return handle_;
        }

        [[nodiscard]] HANDLE release() noexcept
        {
            return std::exchange(handle_, nullptr);
        }

        void reset(HANDLE handle = nullptr) noexcept
        {
            if (handle_ != nullptr && handle_ != INVALID_HANDLE_VALUE)
            {
                ::CloseHandle(handle_);
            }

            handle_ = handle;
        }

    private:
        HANDLE handle_{nullptr};
    };

    [[nodiscard]] std::wstring QuoteCommandLineArgument(std::wstring_view argument)
    {
        if (argument.empty())
        {
            return L"\"\"";
        }

        const bool requires_quotes = argument.find_first_of(L" \t\n\v\"") != std::wstring_view::npos;

        if (!requires_quotes)
        {
            return std::wstring(argument);
        }

        std::wstring quoted;
        quoted.push_back(L'"');

        std::size_t backslash_count = 0;

        for (wchar_t character : argument)
        {
            if (character == L'\\')
            {
                ++backslash_count;
                continue;
            }

            if (character == L'"')
            {
                quoted.append(backslash_count * 2 + 1, L'\\');
                quoted.push_back(L'"');
                backslash_count = 0;
                continue;
            }

            if (backslash_count != 0)
            {
                quoted.append(backslash_count, L'\\');
                backslash_count = 0;
            }

            quoted.push_back(character);
        }

        if (backslash_count != 0)
        {
            quoted.append(backslash_count * 2, L'\\');
        }

        quoted.push_back(L'"');
        return quoted;
    }

    [[nodiscard]] std::wstring BuildCommandLine(
        const std::filesystem::path& executable_path,
        const std::vector<std::wstring>& arguments)
    {
        std::wstring command_line = QuoteCommandLineArgument(executable_path.wstring());

        for (const std::wstring& argument : arguments)
        {
            command_line.push_back(L' ');
            command_line.append(QuoteCommandLineArgument(argument));
        }

        return command_line;
    }

    [[nodiscard]] std::runtime_error MakeProcessError(std::wstring_view message)
    {
        return std::runtime_error(pptxvp::helper::WideToUtf8(message));
    }

    [[nodiscard]] std::wstring FormatLastError(DWORD error_code)
    {
        LPWSTR buffer = nullptr;

        const DWORD flags = FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS;
        const DWORD length = ::FormatMessageW(
            flags,
            nullptr,
            error_code,
            MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
            reinterpret_cast<LPWSTR>(&buffer),
            0,
            nullptr);

        std::wstring message;

        if (length != 0 && buffer != nullptr)
        {
            message.assign(buffer, buffer + length);
            while (!message.empty() && std::iswspace(message.back()))
            {
                message.pop_back();
            }
        }
        else
        {
            message = L"未知的 Win32 错误。";
        }

        if (buffer != nullptr)
        {
            ::LocalFree(buffer);
        }

        return message;
    }
}

namespace pptxvp::helper
{
    namespace
    {
        [[nodiscard]] ProcessResult RunProcessCore(
            const std::filesystem::path& executable_path,
            const std::vector<std::wstring>& arguments,
            const OutputChunkCallback* output_callback,
            const std::filesystem::path& working_directory)
        {
            SECURITY_ATTRIBUTES security_attributes{};
            security_attributes.nLength = sizeof(security_attributes);
            security_attributes.bInheritHandle = TRUE;

            HANDLE read_pipe = nullptr;
            HANDLE write_pipe = nullptr;

            if (!::CreatePipe(&read_pipe, &write_pipe, &security_attributes, 0))
            {
                throw MakeProcessError(L"无法为子进程输出创建管道。");
            }

            UniqueHandle read_handle(read_pipe);
            UniqueHandle write_handle(write_pipe);

            if (!::SetHandleInformation(read_handle.get(), HANDLE_FLAG_INHERIT, 0))
            {
                throw MakeProcessError(L"无法配置进程输出重定向。");
            }

            STARTUPINFOW startup_info{};
            startup_info.cb = sizeof(startup_info);
            startup_info.dwFlags = STARTF_USESTDHANDLES;
            startup_info.hStdInput = ::GetStdHandle(STD_INPUT_HANDLE);
            startup_info.hStdOutput = write_handle.get();
            startup_info.hStdError = write_handle.get();

            PROCESS_INFORMATION process_information{};

            std::wstring command_line = BuildCommandLine(executable_path, arguments);

            if (!::CreateProcessW(
                    executable_path.c_str(),
                    command_line.data(),
                    nullptr,
                    nullptr,
                    TRUE,
                    CREATE_NO_WINDOW,
                    nullptr,
                    working_directory.empty() ? nullptr : working_directory.c_str(),
                    &startup_info,
                    &process_information))
            {
                const DWORD error_code = ::GetLastError();
                throw MakeProcessError(
                    L"无法启动进程： " + executable_path.wstring() + L"。 " + FormatLastError(error_code));
            }

            UniqueHandle process_handle(process_information.hProcess);
            UniqueHandle thread_handle(process_information.hThread);

            write_handle.reset();

            std::string captured_output;
            std::array<char, 4096> buffer{};

            while (true)
            {
                DWORD bytes_read = 0;
                const BOOL read_success =
                    ::ReadFile(read_handle.get(), buffer.data(), static_cast<DWORD>(buffer.size()), &bytes_read, nullptr);

                if (!read_success || bytes_read == 0)
                {
                    break;
                }

                captured_output.append(buffer.data(), buffer.data() + bytes_read);

                if (output_callback != nullptr && *output_callback)
                {
                    (*output_callback)(std::string_view(buffer.data(), bytes_read));
                }
            }

            ::WaitForSingleObject(process_handle.get(), INFINITE);

            DWORD exit_code = 0;
            if (!::GetExitCodeProcess(process_handle.get(), &exit_code))
            {
                throw MakeProcessError(L"无法获取子进程退出代码。");
            }

            return ProcessResult{
                .exit_code = static_cast<int>(exit_code),
                .output = std::move(captured_output),
            };
        }
    }

    ProcessResult RunProcess(
        const std::filesystem::path& executable_path,
        const std::vector<std::wstring>& arguments,
        const std::filesystem::path& working_directory)
    {
        return RunProcessCore(executable_path, arguments, nullptr, working_directory);
    }

    ProcessResult RunProcessStreaming(
        const std::filesystem::path& executable_path,
        const std::vector<std::wstring>& arguments,
        const OutputChunkCallback& output_callback,
        const std::filesystem::path& working_directory)
    {
        return RunProcessCore(executable_path, arguments, &output_callback, working_directory);
    }
}
