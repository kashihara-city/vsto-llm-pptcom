namespace Pptcom.Agent
{
    // A model the endpoint offers (OpenAiCompatibleClient.ListModelsAsync): its mode when known ("chat",
    // "image_generation", ...) and how many tokens it takes in, when the endpoint says.
    internal sealed class ModelInfo(string name, string? mode, int? maxInputTokens)
    {
        public string Name { get; } = name;

        public string? Mode { get; } = mode;

        public int? MaxInputTokens { get; } = maxInputTokens;
    }
}
