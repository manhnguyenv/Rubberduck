using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    [SuppressMessage("ReSharper", "UseIndexedProperty")]
    public class Property : SafeComWrapper<Microsoft.Vbe.Interop.Property>, IProperty
    {
        public Property(Microsoft.Vbe.Interop.Property comObject) 
            : base(comObject)
        {
        }

        public string Name
        {
            get { return IsWrappingNullReference ? string.Empty : ComObject.Name; }
        }

        public int IndexCount
        {
            get { return IsWrappingNullReference ? 0 : ComObject.NumIndices; }
        }

        public IProperties Collection
        {
            get { return new Properties(IsWrappingNullReference ? null : ComObject.Collection); }
        }

        public IProperties Parent
        {
            get { return new Properties(IsWrappingNullReference ? null : ComObject.Parent); }
        }

        public IApplication Application
        {
            get { return new Application(IsWrappingNullReference ? null : ComObject.Application); }
        }

        public IVBE VBE
        {
            get { return new VBE(IsWrappingNullReference ? null : ComObject.VBE); }
        }

        public object Value
        {
            get { return IsWrappingNullReference ? null : ComObject.Value; }
            set { ComObject.Value = value; }
        }

        /// <summary>
        /// Getter can return an unwrapped COM object; remember to call Marshal.ReleaseComObject on the returned object.
        /// </summary>
        public object GetIndexedValue(object index1, object index2 = null, object index3 = null, object index4 = null)
        {
            return ComObject.get_IndexedValue(index1, index2, index3, index4);
        }

        public void SetIndexedValue(object value, object index1, object index2 = null, object index3 = null, object index4 = null)
        {
            ComObject.set_IndexedValue(index1, index2, index3, index4, value);
        }

        /// <summary>
        /// Getter returns an unwrapped COM object; remember to call Marshal.ReleaseComObject on the returned object.
        /// </summary>
        public object Object
        {
            get { return IsWrappingNullReference ? null : ComObject.Object; }
            set { ComObject.Object = value; }
        }

        public override void Release()
        {
            if (!IsWrappingNullReference)
            {
                Marshal.ReleaseComObject(ComObject);
            } 
        }

        public override bool Equals(SafeComWrapper<Microsoft.Vbe.Interop.Property> other)
        {
            return IsEqualIfNull(other) ||
                (other != null && other.ComObject.Name == Name && ReferenceEquals(other.ComObject.Parent, ComObject.Parent));
        }

        public bool Equals(IProperty other)
        {
            return Equals(other as SafeComWrapper<Microsoft.Vbe.Interop.Property>);
        }

        public override int GetHashCode()
        {
            return ComputeHashCode(Name, IndexCount, Parent.ComObject);
        }
    }
}