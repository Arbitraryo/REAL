#include "MinimumLatencyAudioClient.h"

#include <Audioclient.h>
#include <mmdeviceapi.h>
#include <cassert>
#include <wrl/client.h> // For Microsoft::WRL::ComPtr
#include <avrt.h>       // For AvSetMmThreadCharacteristics

#pragma comment(lib, "Avrt.lib")

using namespace miniant::Windows;
using namespace miniant::Windows::WasapiLatency;

const CLSID CLSID_MMDeviceEnumerator = __uuidof(MMDeviceEnumerator);
const IID IID_IMMDeviceEnumerator = __uuidof(IMMDeviceEnumerator);
const IID IID_IAudioClient3 = __uuidof(IAudioClient3);

#define CHECK_HR(hr) if (FAILED(hr)) return tl::make_unexpected(WindowsError());

MinimumLatencyAudioClient::MinimumLatencyAudioClient(MinimumLatencyAudioClient&& other) {
    assert(other.m_pAudioClient != nullptr);
    assert(other.m_pFormat != nullptr);

    m_pAudioClient = other.m_pAudioClient;
    m_pFormat = other.m_pFormat;

    other.m_pAudioClient = nullptr;
    other.m_pFormat = nullptr;
}

MinimumLatencyAudioClient::MinimumLatencyAudioClient(void* pAudioClient, void* pFormat) :
    m_pAudioClient(pAudioClient), m_pFormat(pFormat) {}

MinimumLatencyAudioClient::~MinimumLatencyAudioClient() {
    Uninitialise();
}

MinimumLatencyAudioClient& MinimumLatencyAudioClient::operator= (MinimumLatencyAudioClient&& rhs) {
    assert(rhs.m_pAudioClient != nullptr);
    assert(rhs.m_pFormat != nullptr);

    Uninitialise();
    m_pAudioClient = rhs.m_pAudioClient;
    m_pFormat = rhs.m_pFormat;

    rhs.m_pAudioClient = nullptr;
    rhs.m_pFormat = nullptr;

    return *this;
}

void MinimumLatencyAudioClient::Uninitialise() {
    if (m_pAudioClient == nullptr) {
        assert(m_pFormat == nullptr);
        return;
    }

    assert(m_pFormat != nullptr);

    static_cast<IAudioClient3*>(m_pAudioClient)->Release();
    m_pAudioClient = nullptr;

    CoTaskMemFree(m_pFormat);
    m_pFormat = nullptr;
}

tl::expected<MinimumLatencyAudioClient::Properties, WindowsError> MinimumLatencyAudioClient::GetProperties() {
    Properties properties;
    HRESULT hr = static_cast<IAudioClient3*>(m_pAudioClient)->GetSharedModeEnginePeriod(
        static_cast<WAVEFORMATEX*>(m_pFormat),
        &properties.defaultBufferSize,
        &properties.fundamentalBufferSize,
        &properties.minimumBufferSize,
        &properties.maximumBufferSize);
    if (FAILED(hr)) {
        return tl::make_unexpected(WindowsError());
    }

    properties.sampleRate = static_cast<WAVEFORMATEX*>(m_pFormat)->nSamplesPerSec;
    properties.bitsPerSample = static_cast<WAVEFORMATEX*>(m_pFormat)->wBitsPerSample;
    properties.numChannels = static_cast<WAVEFORMATEX*>(m_pFormat)->nChannels;

    return properties;
}

UINT32 CalculateBufferSize(UINT32 minPeriod, UINT32 maxPeriod) {
    return std::max(minPeriod, std::min(maxPeriod, 512)); // Minimum 512 frames
}

tl::expected<MinimumLatencyAudioClient, WindowsError> MinimumLatencyAudioClient::Start() {
    HRESULT hr;

    Microsoft::WRL::ComPtr<IMMDeviceEnumerator> pEnumerator;
    Microsoft::WRL::ComPtr<IMMDevice> pDevice;
    Microsoft::WRL::ComPtr<IAudioClient3> pAudioClient;
    WAVEFORMATEX* pFormat = nullptr;

    // Initialize COM
    hr = CoInitialize(NULL);
    CHECK_HR(hr);

    // Create device enumerator
    hr = CoCreateInstance(
        CLSID_MMDeviceEnumerator,
        NULL,
        CLSCTX_ALL,
        IID_IMMDeviceEnumerator,
        reinterpret_cast<void**>(pEnumerator.GetAddressOf()));
    CHECK_HR(hr);

    // Get default audio endpoint
    hr = pEnumerator->GetDefaultAudioEndpoint(eRender, eConsole, pDevice.GetAddressOf());
    CHECK_HR(hr);

    // Activate audio client
    hr = pDevice->Activate(IID_IAudioClient3, CLSCTX_ALL, NULL, reinterpret_cast<void**>(pAudioClient.GetAddressOf()));
    CHECK_HR(hr);

    // Get mix format
    hr = pAudioClient->GetMixFormat(&pFormat);
    CHECK_HR(hr);

    // Get shared mode engine period
    UINT32 defaultPeriodInFrames;
    UINT32 fundamentalPeriodInFrames;
    UINT32 minPeriodInFrames;
    UINT32 maxPeriodInFrames;
    hr = pAudioClient->GetSharedModeEnginePeriod(
        pFormat,
        &defaultPeriodInFrames,
        &fundamentalPeriodInFrames,
        &minPeriodInFrames,
        &maxPeriodInFrames);
    CHECK_HR(hr);

    // Calculate optimal buffer size
    UINT32 optimalBufferSize = CalculateBufferSize(minPeriodInFrames, maxPeriodInFrames);

    // Initialize audio stream
    hr = pAudioClient->InitializeSharedAudioStream(
        0,
        optimalBufferSize,
        pFormat,
        NULL);
    CHECK_HR(hr);

    // Set thread priority
    DWORD taskIndex = 0;
    HANDLE hTask = AvSetMmThreadCharacteristics(L"Pro Audio", &taskIndex);
    if (!hTask) {
        return tl::make_unexpected(WindowsError());
    }

    // Start audio client
    hr = pAudioClient->Start();
    CHECK_HR(hr);

    return MinimumLatencyAudioClient(pAudioClient.Detach(), pFormat);
}
