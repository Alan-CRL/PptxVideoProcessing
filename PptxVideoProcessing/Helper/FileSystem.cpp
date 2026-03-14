module;

#define NOMINMAX
#include <Windows.h>

module PptxVideoProcessing.Helper.FileSystem;

import std.compat;

import PptxVideoProcessing.Helper.Utf;

namespace
{
    [[nodiscard]] std::runtime_error MakeFilesystemError(std::wstring_view message)
    {
        return std::runtime_error(pptxvp::helper::WideToUtf8(message));
    }

    [[nodiscard]] std::wstring MakeUniqueSuffix(std::uint64_t high, std::uint64_t low)
    {
        std::wostringstream stream;
        stream << std::hex << high << low;
        return stream.str();
    }

    [[nodiscard]] std::wstring AppendIndex(std::wstring_view base_name, std::size_t index)
    {
        std::wostringstream stream;
        stream << base_name << L'(' << index << L')';
        return stream.str();
    }
}

namespace pptxvp::helper
{
    ScopedTempDirectory::ScopedTempDirectory(std::filesystem::path path) : path_(std::move(path))
    {
    }

    ScopedTempDirectory::ScopedTempDirectory(ScopedTempDirectory&& other) noexcept : path_(std::move(other.path_))
    {
        other.path_.clear();
    }

    ScopedTempDirectory& ScopedTempDirectory::operator=(ScopedTempDirectory&& other) noexcept
    {
        if (this != &other)
        {
            reset();
            path_ = std::move(other.path_);
            other.path_.clear();
        }

        return *this;
    }

    ScopedTempDirectory::~ScopedTempDirectory()
    {
        reset();
    }

    const std::filesystem::path& ScopedTempDirectory::path() const noexcept
    {
        return path_;
    }

    bool ScopedTempDirectory::empty() const noexcept
    {
        return path_.empty();
    }

    void ScopedTempDirectory::reset()
    {
        if (path_.empty())
        {
            return;
        }

        std::error_code error_code;
        std::filesystem::remove_all(path_, error_code);
        path_.clear();
    }

    std::filesystem::path GetExecutablePath()
    {
        std::wstring buffer(512, L'\0');

        while (true)
        {
            DWORD copied_length = ::GetModuleFileNameW(nullptr, buffer.data(), static_cast<DWORD>(buffer.size()));

            if (copied_length == 0)
            {
                throw MakeFilesystemError(L"无法定位当前可执行文件路径。");
            }

            if (copied_length < buffer.size() - 1)
            {
                buffer.resize(copied_length);
                return std::filesystem::path(buffer);
            }

            buffer.resize(buffer.size() * 2);
        }
    }

    std::filesystem::path GetExecutableDirectory()
    {
        return GetExecutablePath().parent_path();
    }

    ScopedTempDirectory CreateUniqueTempDirectory(std::wstring_view prefix)
    {
        const std::filesystem::path root = std::filesystem::temp_directory_path() / L"PptxVideoProcessing";
        std::filesystem::create_directories(root);

        std::random_device random_device;
        std::mt19937_64 generator(random_device());

        for (std::size_t attempt = 0; attempt < 128; ++attempt)
        {
            const std::wstring suffix = MakeUniqueSuffix(generator(), generator());
            const std::filesystem::path candidate = root / std::filesystem::path(std::wstring(prefix) + L"-" + suffix);

            std::error_code error_code;
            if (std::filesystem::create_directories(candidate, error_code))
            {
                return ScopedTempDirectory(candidate);
            }
        }

        throw MakeFilesystemError(L"无法创建唯一的临时工作目录。");
    }

    std::string ReadTextFileUtf8(const std::filesystem::path& path)
    {
        std::ifstream input(path, std::ios::binary);

        if (!input)
        {
            throw MakeFilesystemError(L"无法打开文件： " + path.wstring());
        }

        std::ostringstream stream;
        stream << input.rdbuf();

        const std::string content = stream.str();

        return StripUtf8Bom(content);
    }

    void WriteTextFileUtf8(const std::filesystem::path& path, std::string_view text)
    {
        std::filesystem::create_directories(path.parent_path());

        std::ofstream output(path, std::ios::binary | std::ios::trunc);

        if (!output)
        {
            throw MakeFilesystemError(L"无法写入文件： " + path.wstring());
        }

        output.write(text.data(), static_cast<std::streamsize>(text.size()));

        if (!output)
        {
            throw MakeFilesystemError(L"无法写入文件： " + path.wstring());
        }
    }

    std::filesystem::path MakeUniqueProcessedOutputPath(const std::filesystem::path& input_path)
    {
        const std::wstring stem = input_path.stem().wstring();
        const std::wstring extension = input_path.extension().wstring();
        const std::wstring processed_suffix = L"_\u5df2\u5904\u7406";

        const std::filesystem::path desired_path =
            input_path.parent_path() / std::filesystem::path(stem + processed_suffix + extension);

        if (!std::filesystem::exists(desired_path))
        {
            return desired_path;
        }

        for (std::size_t index = 1; index < 10'000; ++index)
        {
            const std::filesystem::path candidate =
                input_path.parent_path() /
                std::filesystem::path(stem + processed_suffix + L"(" + std::to_wstring(index) + L")" + extension);

            if (!std::filesystem::exists(candidate))
            {
                return candidate;
            }
        }

        throw MakeFilesystemError(L"无法在源演示文稿文件旁创建唯一的输出文件名。");
    }

    std::filesystem::path MakeUniqueSiblingPath(const std::filesystem::path& desired_path)
    {
        if (!std::filesystem::exists(desired_path))
        {
            return desired_path;
        }

        const std::filesystem::path parent = desired_path.parent_path();
        const std::wstring stem = desired_path.stem().wstring();
        const std::wstring extension = desired_path.extension().wstring();

        for (std::size_t index = 1; index < 10'000; ++index)
        {
            const std::filesystem::path candidate =
                parent / std::filesystem::path(stem + L"(" + std::to_wstring(index) + L")" + extension);

            if (!std::filesystem::exists(candidate))
            {
                return candidate;
            }
        }

        throw MakeFilesystemError(L"无法创建唯一的文件路径。");
    }
}
