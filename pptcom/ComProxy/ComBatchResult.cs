using System.Collections.Generic;

namespace Pptcom.ComProxy
{
    internal sealed class ComBatchResult
    {
        public bool Success;
        public readonly List<ComOpResult> Ops = [];
    }
}
