using System;
using System.Collections;
using System.Collections.Generic;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using VB = Microsoft.VB6.Interop.VBIDE;

namespace Rubberduck.VBEditor.SafeComWrappers.VB6
{
    public class References : SafeComWrapper<VB.References>, IReferences
    {
        public References(VB.References target, bool rewrapping = false) 
            : base(target, rewrapping)
        {
            if (!IsWrappingNullReference)
            {
                target.ItemAdded += Target_ItemAdded;
                target.ItemRemoved += Target_ItemRemoved;
            }
        }

        public event EventHandler<ReferenceEventArgs> ItemAdded;
        public event EventHandler<ReferenceEventArgs> ItemRemoved;

        public bool EnableEvents { get; set; }

        public int Count => IsWrappingNullReference ? 0 : Target.Count;

        public IVBProject Parent => new VBProject(IsWrappingNullReference ? null : Target.Parent);

        public IVBE VBE => new VBE(IsWrappingNullReference ? null : Target.VBE);

        private void Target_ItemRemoved(VB.Reference reference)
        {
            if (!EnableEvents) { return; }
            var referenceWrapper = new Reference(reference);
            var handler = ItemRemoved;
            if (handler == null)
            {
                referenceWrapper.Dispose();
                return;
            }
            handler.Invoke(this, new ReferenceEventArgs(referenceWrapper));
        }

        private void Target_ItemAdded(VB.Reference reference)
        {
            if (!EnableEvents) { return; }
            var referenceWrapper = new Reference(reference);
            var handler = ItemAdded;
            if (handler == null)
            {
                referenceWrapper.Dispose();
                return;
            }
            handler.Invoke(this, new ReferenceEventArgs(referenceWrapper));
        }

        public IReference this[object index] => new Reference(Target.Item(index));

        public IReference AddFromGuid(string guid, int major, int minor)
        {
            return new Reference(Target.AddFromGuid(guid, major, minor));
        }

        public IReference AddFromFile(string path)
        {
            return new Reference(Target.AddFromFile(path));
        }

        public void Remove(IReference reference)
        {
            Target.Remove(((ISafeComWrapper<VB.Reference>)reference).Target);
        }

        IEnumerator<IReference> IEnumerable<IReference>.GetEnumerator()
        {
            return new ComWrapperEnumerator<IReference>(Target, comObject => new Reference((VB.Reference)comObject));
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return ((IEnumerable<IReference>)this).GetEnumerator();
        }

        //public override void Release(bool final = false)
        //{
        //    if (!IsWrappingNullReference)
        //    {
        //        Target.ItemAdded -= Target_ItemAdded;
        //        Target.ItemRemoved -= Target_ItemRemoved;
        //        for (var i = 1; i <= Count; i++)
        //        {
        //            this[i].Release();
        //        }
        //        base.Release(final);
        //    }
        //}

        public override bool Equals(ISafeComWrapper<VB.References> other)
        {
            return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.Target.Parent, Parent.Target));
        }

        public bool Equals(IReferences other)
        {
            return Equals(other as SafeComWrapper<VB.References>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : Target.GetHashCode();
        }
    }
}