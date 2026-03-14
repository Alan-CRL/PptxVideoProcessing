module;

module PptxVideoProcessing.OfficeXml;

import std.compat;

import PptxVideoProcessing.Helper.FileSystem;
import PptxVideoProcessing.Helper.Utf;
import PptxVideoProcessing.Media;

namespace
{
    [[nodiscard]] std::runtime_error MakeOfficeXmlError(std::wstring_view message)
    {
        return std::runtime_error(pptxvp::helper::WideToUtf8(message));
    }

    [[nodiscard]] std::size_t ReplaceRelationshipTargets(
        std::string& xml_text,
        const std::vector<pptxvp::MediaRename>& renames)
    {
        std::size_t replacements = 0;
        std::size_t search_position = 0;

        while (true)
        {
            const std::size_t target_position = xml_text.find("Target=\"", search_position);

            if (target_position == std::string::npos)
            {
                break;
            }

            const std::size_t value_begin = target_position + 8;
            const std::size_t value_end = xml_text.find('"', value_begin);

            if (value_end == std::string::npos)
            {
                break;
            }

            const std::string current_target = xml_text.substr(value_begin, value_end - value_begin);
            const std::size_t separator_position = current_target.find_last_of("/\\");
            const std::string current_leaf =
                separator_position == std::string::npos ? current_target : current_target.substr(separator_position + 1);

            bool replaced = false;

            for (const pptxvp::MediaRename& rename : renames)
            {
                const std::string old_leaf = pptxvp::helper::WideToUtf8(rename.original_relative_path.filename().wstring());
                const std::string new_leaf = pptxvp::helper::WideToUtf8(rename.updated_relative_path.filename().wstring());

                if (current_leaf != old_leaf)
                {
                    continue;
                }

                const std::string updated_target =
                    separator_position == std::string::npos
                        ? new_leaf
                        : current_target.substr(0, separator_position + 1) + new_leaf;

                xml_text.replace(value_begin, value_end - value_begin, updated_target);
                search_position = value_begin + updated_target.size();
                ++replacements;
                replaced = true;
                break;
            }

            if (!replaced)
            {
                search_position = value_end + 1;
            }
        }

        return replacements;
    }

    [[nodiscard]] bool NeedsMp4ContentType(const std::vector<pptxvp::MediaRename>& renames)
    {
        return std::ranges::any_of(renames, [](const pptxvp::MediaRename& rename)
        {
            return pptxvp::helper::ToLowerAscii(rename.updated_relative_path.extension().wstring()) == L".mp4";
        });
    }

    [[nodiscard]] std::size_t EnsureMp4ContentType(const std::filesystem::path& content_types_path)
    {
        std::string xml_text = pptxvp::helper::ReadTextFileUtf8(content_types_path);

        if (xml_text.find("Extension=\"mp4\"") != std::string::npos ||
            xml_text.find("Extension='mp4'") != std::string::npos)
        {
            return 0;
        }

        const std::string insertion = "  <Default Extension=\"mp4\" ContentType=\"video/mp4\"/>\r\n";
        const std::size_t closing_tag = xml_text.rfind("</Types>");

        if (closing_tag == std::string::npos)
        {
            throw MakeOfficeXmlError(L"在 [Content_Types].xml 中找不到结束标签 </Types>。");
        }

        xml_text.insert(closing_tag, insertion);
        pptxvp::helper::WriteTextFileUtf8(content_types_path, xml_text);
        return 1;
    }
}

namespace pptxvp
{
    std::size_t UpdateOfficeMediaReferences(
        const std::filesystem::path& extracted_root,
        const std::vector<MediaRename>& renames)
    {
        std::size_t updated_items = 0;

        for (const auto& entry : std::filesystem::recursive_directory_iterator(extracted_root))
        {
            if (!entry.is_regular_file())
            {
                continue;
            }

            if (helper::ToLowerAscii(entry.path().extension().wstring()) != L".rels")
            {
                continue;
            }

            std::string xml_text = helper::ReadTextFileUtf8(entry.path());
            const std::size_t replacements = ReplaceRelationshipTargets(xml_text, renames);

            if (replacements != 0)
            {
                helper::WriteTextFileUtf8(entry.path(), xml_text);
                updated_items += replacements;
            }
        }

        if (NeedsMp4ContentType(renames))
        {
            updated_items += EnsureMp4ContentType(extracted_root / "[Content_Types].xml");
        }

        return updated_items;
    }
}
