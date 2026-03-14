module;

#define NOMINMAX
#include <Windows.h>
#include <objbase.h>
#include <shobjidl.h>

module PptxVideoProcessing.Ui;

import std.compat;

import PptxVideoProcessing.Helper.Utf;

namespace
{
    template <typename T>
    class ComPtr
    {
    public:
        ComPtr() = default;
        ComPtr(const ComPtr&) = delete;
        ComPtr& operator=(const ComPtr&) = delete;

        ComPtr(ComPtr&& other) noexcept : pointer_(std::exchange(other.pointer_, nullptr))
        {
        }

        ComPtr& operator=(ComPtr&& other) noexcept
        {
            if (this != &other)
            {
                reset();
                pointer_ = std::exchange(other.pointer_, nullptr);
            }

            return *this;
        }

        ~ComPtr()
        {
            reset();
        }

        [[nodiscard]] T* get() const noexcept
        {
            return pointer_;
        }

        [[nodiscard]] T** put() noexcept
        {
            reset();
            return &pointer_;
        }

        [[nodiscard]] T* operator->() const noexcept
        {
            return pointer_;
        }

        void reset(T* pointer = nullptr) noexcept
        {
            if (pointer_ != nullptr)
            {
                pointer_->Release();
            }

            pointer_ = pointer;
        }

    private:
        T* pointer_{nullptr};
    };

    class ComApartment
    {
    public:
        ComApartment()
        {
            const HRESULT result = ::CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
            initialized_ = SUCCEEDED(result);

            if (!initialized_ && result != RPC_E_CHANGED_MODE)
            {
                std::wostringstream stream;
                stream << L"初始化文件选择对话框所需的 COM 环境失败。HRESULT=0x"
                       << std::hex << static_cast<unsigned long>(result);
                throw std::runtime_error(pptxvp::helper::WideToUtf8(stream.str()));
            }
        }

        ~ComApartment()
        {
            if (initialized_)
            {
                ::CoUninitialize();
            }
        }

    private:
        bool initialized_{false};
    };

    class CoTaskMemString
    {
    public:
        CoTaskMemString() = default;
        CoTaskMemString(const CoTaskMemString&) = delete;
        CoTaskMemString& operator=(const CoTaskMemString&) = delete;

        ~CoTaskMemString()
        {
            if (value_ != nullptr)
            {
                ::CoTaskMemFree(value_);
            }
        }

        [[nodiscard]] PWSTR* put() noexcept
        {
            if (value_ != nullptr)
            {
                ::CoTaskMemFree(value_);
                value_ = nullptr;
            }

            return &value_;
        }

        [[nodiscard]] std::wstring get() const
        {
            return value_ == nullptr ? std::wstring{} : std::wstring(value_);
        }

    private:
        PWSTR value_{nullptr};
    };
}

namespace pptxvp
{
    std::filesystem::path PickPowerPointFile()
    {
        ComApartment com_apartment;

        ComPtr<IFileOpenDialog> dialog;
        const HRESULT create_result = ::CoCreateInstance(
            CLSID_FileOpenDialog,
            nullptr,
            CLSCTX_INPROC_SERVER,
            IID_PPV_ARGS(dialog.put()));

        if (FAILED(create_result))
        {
            std::wostringstream stream;
            stream << L"创建文件选择对话框失败。HRESULT=0x"
                   << std::hex << static_cast<unsigned long>(create_result);
            throw std::runtime_error(helper::WideToUtf8(stream.str()));
        }

        DWORD options = 0;
        dialog->GetOptions(&options);
        dialog->SetOptions(options | FOS_FORCEFILESYSTEM | FOS_FILEMUSTEXIST | FOS_PATHMUSTEXIST);

        const COMDLG_FILTERSPEC filters[] = {
            {L"PPTX 演示文稿 (*.pptx)", L"*.pptx"},
        };

        dialog->SetFileTypes(static_cast<UINT>(std::size(filters)), filters);
        dialog->SetDefaultExtension(L"pptx");
        dialog->SetTitle(L"请选择要处理的 PPTX 文件");

        const HRESULT show_result = dialog->Show(nullptr);

        if (show_result == HRESULT_FROM_WIN32(ERROR_CANCELLED))
        {
            return {};
        }

        if (FAILED(show_result))
        {
            std::wostringstream stream;
            stream << L"显示文件选择对话框失败。HRESULT=0x"
                   << std::hex << static_cast<unsigned long>(show_result);
            throw std::runtime_error(helper::WideToUtf8(stream.str()));
        }

        ComPtr<IShellItem> item;
        const HRESULT get_result = dialog->GetResult(item.put());

        if (FAILED(get_result))
        {
            std::wostringstream stream;
            stream << L"读取已选文件失败。HRESULT=0x"
                   << std::hex << static_cast<unsigned long>(get_result);
            throw std::runtime_error(helper::WideToUtf8(stream.str()));
        }

        CoTaskMemString selected_path;
        const HRESULT display_name_result = item->GetDisplayName(SIGDN_FILESYSPATH, selected_path.put());

        if (FAILED(display_name_result))
        {
            std::wostringstream stream;
            stream << L"读取所选文件路径失败。HRESULT=0x"
                   << std::hex << static_cast<unsigned long>(display_name_result);
            throw std::runtime_error(helper::WideToUtf8(stream.str()));
        }

        return std::filesystem::path(selected_path.get());
    }
}
