/*
PptxVideoProcessing 使用说明

一、程序用途
1. 本程序用于处理 .pptx 演示文稿中的内嵌视频。
2. 程序会解包 PPTX，查找 ppt/media 目录下的常见视频文件，并调用 ffmpeg.exe 进行转码或调整。
3. 处理完成后，程序会重新打包为新的 .pptx 文件，原文件不会被直接覆盖。

二、使用步骤
1. 将本程序生成的 exe 与 ffmpeg.exe 放在同一目录。
2. 可选：在 exe 同目录放置 config.json；如果不存在，程序首次启动时会自动创建一个空的 {}。
3. 运行程序，按提示选择一个 .pptx 文件。
4. 程序会自动处理内嵌视频，并在原文件所在目录生成“原文件名_已处理.pptx”。
5. 如果输出文件已存在，程序会自动递增命名为“_已处理(2)”“_已处理(3)”等。

三、config.json 示例
{
  "encoder": "h265",
  "frameRate": 30,
  "resolution": "720p"
}

四、config.json 配置项说明
1. encoder
   - 可选。
   - 用于指定视频编码器。
   - 支持别名：
     h264  -> libx264
     h265  -> libx265
     av1   -> libsvtav1
     mpeg4 -> mpeg4
   - 也支持直接填写 ffmpeg.exe 支持的编码器名称，例如：
     libx264、libx265、libsvtav1、mpeg4、
     h264_nvenc、hevc_nvenc、av1_nvenc、
     h264_qsv、hevc_qsv、av1_qsv 等。
   - 如果该项缺省，而 frameRate 或 resolution 有值，程序会尽量探测原视频编码并保持原编码风格。

2. frameRate
   - 可选。
   - 必须为正整数。
   - 常见值：24、25、30、50、60。
   - 缺省表示不修改帧率。

3. resolution
   - 可选。
   - 当前支持：360p、480p、720p、1080p、2160p。
   - 程序会按目标高度缩放，并保持原视频宽高比。
   - 缺省表示不修改分辨率。

五、补充说明
1. 仅支持 .pptx，不支持旧版二进制 .ppt。
2. config.json 支持 UTF-8 和 UTF-8 BOM。
3. 某个配置项缺省时，仅表示不修改该项，不会报错。
4. 程序会扫描 ppt/media 中的常见视频文件，例如 mp4、m4v、mov、avi、wmv、mkv、webm。
5. 若原视频容器与目标编码不兼容，程序会转成 MP4，并自动更新 PPTX 内的媒体引用。
6. 程序默认保留首个视频流和可选音频流，音频通常按复制方式保留，不主动改变演示稿中的媒体位置与时长设置。
*/
import std.compat;

import PptxVideoProcessing.App;
import PptxVideoProcessing.Helper.Console;
import PptxVideoProcessing.Helper.Utf;

int wmain()
{
    pptxvp::helper::InitializeUtf16Console();
    pptxvp::helper::WriteLine(L"版权所有 (c) 2026 AlanCRL(陈润林) 工作室");
    pptxvp::helper::WriteLine(L"本项目基于 GNU 通用公共许可证第 3 版获得许可");
    pptxvp::helper::WriteLine(L"------------------------------------------------");
    pptxvp::helper::WriteLine(L"");

    int exit_code = 0;

    try
    {
        exit_code = pptxvp::Run();
    }
    catch (const std::exception& exception)
    {
        pptxvp::helper::WriteErrorLine(L"错误： " + pptxvp::helper::Utf8ToWide(exception.what()));
        exit_code = 1;
    }
    catch (...)
    {
        pptxvp::helper::WriteErrorLine(L"错误：发生了未预期的异常。");
        exit_code = 1;
    }

    pptxvp::helper::WriteLine(L"");
    pptxvp::helper::WaitForAnyKey();
    return exit_code;
}
