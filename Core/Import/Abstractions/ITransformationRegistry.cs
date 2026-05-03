using SheetSet.Core.Import.Models;

namespace SheetSet.Core.Import.Abstractions;

public interface ITransformationRegistry
{
    void Register(ITransformationRule rule);
    ITransformationRule Resolve(TransformationType type);
    bool IsRegistered(TransformationType type);
}
