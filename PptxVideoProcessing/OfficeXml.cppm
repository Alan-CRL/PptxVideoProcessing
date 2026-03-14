export module PptxVideoProcessing.OfficeXml;

import std.compat;

import PptxVideoProcessing.Media;

export namespace pptxvp
{
    [[nodiscard]] std::size_t UpdateOfficeMediaReferences(
        const std::filesystem::path& extracted_root,
        const std::vector<MediaRename>& renames);
}
