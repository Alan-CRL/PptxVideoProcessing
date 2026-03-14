module;

#define ZIP_STATIC
#include <zip.h>

module PptxVideoProcessing.Archive;

import std.compat;

import PptxVideoProcessing.Helper.Utf;

namespace
{
    [[nodiscard]] std::runtime_error MakeArchiveError(std::wstring_view message)
    {
        return std::runtime_error(pptxvp::helper::WideToUtf8(message));
    }

    [[nodiscard]] std::wstring ZipErrorToWide(zip_error_t& error)
    {
        return pptxvp::helper::Utf8ToWide(zip_error_strerror(&error));
    }

    [[nodiscard]] std::wstring ArchiveErrorMessage(zip_t* archive)
    {
        return pptxvp::helper::Utf8ToWide(zip_strerror(archive));
    }

    [[nodiscard]] std::filesystem::path ValidateArchiveEntryPath(std::string_view entry_name)
    {
        const std::filesystem::path relative_path = std::filesystem::path(std::string(entry_name)).lexically_normal();

        if (relative_path.empty() || relative_path == ".")
        {
            return {};
        }

        if (relative_path.is_absolute())
        {
            throw MakeArchiveError(L"压缩包条目路径是绝对路径，这是不允许的。");
        }

        for (const auto& part : relative_path)
        {
            if (part == "..")
            {
                throw MakeArchiveError(L"压缩包条目路径越出了目标解压目录。");
            }
        }

        return relative_path;
    }

    [[nodiscard]] zip_t* OpenArchiveForReading(const std::filesystem::path& archive_path)
    {
        zip_error_t error;
        zip_error_init(&error);

        zip_source_t* source = zip_source_win32w_create(archive_path.c_str(), 0, ZIP_LENGTH_TO_END, &error);

        if (source == nullptr)
        {
            const std::wstring message = ZipErrorToWide(error);
            zip_error_fini(&error);
            throw MakeArchiveError(L"无法打开压缩包源文件： " + archive_path.wstring() + L"。 " + message);
        }

        zip_t* archive = zip_open_from_source(source, 0, &error);

        if (archive == nullptr)
        {
            const std::wstring message = ZipErrorToWide(error);
            zip_source_free(source);
            zip_error_fini(&error);
            throw MakeArchiveError(L"无法打开 PPTX 压缩包： " + archive_path.wstring() + L"。 " + message);
        }

        zip_error_fini(&error);
        return archive;
    }

    [[nodiscard]] zip_t* OpenArchiveForWriting(const std::filesystem::path& archive_path)
    {
        zip_error_t error;
        zip_error_init(&error);

        zip_source_t* source = zip_source_win32w_create(archive_path.c_str(), 0, ZIP_LENGTH_TO_END, &error);

        if (source == nullptr)
        {
            const std::wstring message = ZipErrorToWide(error);
            zip_error_fini(&error);
            throw MakeArchiveError(L"无法创建压缩包源文件： " + archive_path.wstring() + L"。 " + message);
        }

        zip_t* archive = zip_open_from_source(source, ZIP_CREATE | ZIP_TRUNCATE, &error);

        if (archive == nullptr)
        {
            const std::wstring message = ZipErrorToWide(error);
            zip_source_free(source);
            zip_error_fini(&error);
            throw MakeArchiveError(L"无法创建 PPTX 压缩包： " + archive_path.wstring() + L"。 " + message);
        }

        zip_error_fini(&error);
        return archive;
    }
}

namespace pptxvp
{
    void ExtractArchive(const std::filesystem::path& archive_path, const std::filesystem::path& destination_directory)
    {
        zip_t* archive = OpenArchiveForReading(archive_path);

        try
        {
            const zip_int64_t entry_count = zip_get_num_entries(archive, 0);

            for (zip_uint64_t index = 0; index < static_cast<zip_uint64_t>(entry_count); ++index)
            {
                zip_stat_t stat{};
                zip_stat_init(&stat);

                if (zip_stat_index(archive, index, 0, &stat) != 0)
                {
                    throw MakeArchiveError(L"读取压缩包条目元数据失败。 " + ArchiveErrorMessage(archive));
                }

                const std::string entry_name = stat.name == nullptr ? std::string{} : std::string(stat.name);
                const std::filesystem::path relative_path = ValidateArchiveEntryPath(entry_name);

                if (relative_path.empty())
                {
                    continue;
                }

                const std::filesystem::path output_path = destination_directory / relative_path;

                if (!entry_name.empty() && entry_name.back() == '/')
                {
                    std::filesystem::create_directories(output_path);
                    continue;
                }

                std::filesystem::create_directories(output_path.parent_path());

                zip_file_t* archive_file = zip_fopen_index(archive, index, 0);

                if (archive_file == nullptr)
                {
                    throw MakeArchiveError(
                        L"无法打开待解压的压缩包条目。 " + ArchiveErrorMessage(archive));
                }

                std::ofstream output(output_path, std::ios::binary | std::ios::trunc);

                if (!output)
                {
                    zip_fclose(archive_file);
                    throw MakeArchiveError(L"无法创建解压后的文件： " + output_path.wstring());
                }

                std::array<std::byte, 8192> buffer{};

                while (true)
                {
                    const zip_int64_t bytes_read = zip_fread(archive_file, buffer.data(), buffer.size());

                    if (bytes_read < 0)
                    {
                        zip_fclose(archive_file);
                        throw MakeArchiveError(L"解压压缩包数据失败。 " + ArchiveErrorMessage(archive));
                    }

                    if (bytes_read == 0)
                    {
                        break;
                    }

                    output.write(
                        reinterpret_cast<const char*>(buffer.data()),
                        static_cast<std::streamsize>(bytes_read));

                    if (!output)
                    {
                        zip_fclose(archive_file);
                        throw MakeArchiveError(L"写入解压后的文件失败： " + output_path.wstring());
                    }
                }

                zip_fclose(archive_file);
            }
        }
        catch (...)
        {
            zip_discard(archive);
            throw;
        }

        if (zip_close(archive) != 0)
        {
            throw MakeArchiveError(L"关闭已打开的 PPTX 压缩包时失败。");
        }
    }

    void CreateArchiveFromDirectory(const std::filesystem::path& source_directory, const std::filesystem::path& archive_path)
    {
        if (!archive_path.parent_path().empty())
        {
            std::filesystem::create_directories(archive_path.parent_path());
        }

        zip_t* archive = OpenArchiveForWriting(archive_path);

        try
        {
            std::vector<std::filesystem::path> files;

            for (const auto& entry : std::filesystem::recursive_directory_iterator(source_directory))
            {
                if (entry.is_regular_file())
                {
                    files.push_back(entry.path());
                }
            }

            std::ranges::sort(files);

            for (const std::filesystem::path& file_path : files)
            {
                const std::filesystem::path relative_path = std::filesystem::relative(file_path, source_directory);
                const std::string archive_name = relative_path.generic_string();

                zip_source_t* file_source = zip_source_win32w(archive, file_path.c_str(), 0, ZIP_LENGTH_TO_END);

                if (file_source == nullptr)
                {
                    throw MakeArchiveError(L"无法为文件创建压缩包源： " + file_path.wstring());
                }

                if (zip_file_add(archive, archive_name.c_str(), file_source, ZIP_FL_ENC_UTF_8) < 0)
                {
                    zip_source_free(file_source);
                    throw MakeArchiveError(
                        L"无法将文件写入输出 PPTX 压缩包。 " + ArchiveErrorMessage(archive));
                }
            }
        }
        catch (...)
        {
            zip_discard(archive);
            throw;
        }

        if (zip_close(archive) != 0)
        {
            throw MakeArchiveError(L"输出 PPTX 压缩包写入完成时失败。");
        }
    }
}
