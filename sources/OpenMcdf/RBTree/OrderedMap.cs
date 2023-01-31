using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.CompilerServices;

namespace OpenMcdf.RBTree;

#pragma warning disable SA1009 // Closing parenthesis should be spaced correctly
#pragma warning disable SA1124 // Do not use regions
#pragma warning disable SA1202 // Elements should be ordered by access
#pragma warning disable SA1602 // Enumeration items should be documented

/// <summary>
/// Color of a node in a Red-Black tree.
/// </summary>
internal enum NodeColor : byte
{
    Red,
    Black,
    Unused,
    LinkedList,
}

/// <summary>
/// Represents a collection of objects that is maintained in sorted order. <see cref="OrderedMap{TKey, TValue}"/> uses Red-Black Tree structure to store objects.
/// </summary>
/// <typeparam name="TKey">The type of keys in the collection.</typeparam>
/// <typeparam name="TValue">The type of values in the collection.</typeparam>
public class OrderedMap<TKey, TValue> : IDictionary<TKey, TValue>, IReadOnlyDictionary<TKey, TValue>, IDictionary where TValue : class, IRbNode
{
    #region Node

    // internal Func<TKey, TValue, NodeColor, Node> CreateNode { get; set; } = static (key, value, color) => new Node(key, value, color);

    /// <summary>
    /// Represents a node in a <see cref="OrderedMap{TKey, TValue}"/>.
    /// </summary>
    public class Node
    {
        

        internal Node(TKey key, TValue value, NodeColor color)
        {
            Key = key;
            Value = value;
            Color = color;
        }

        /// <summary>
        /// Gets the key contained in the node.
        /// </summary>
        public TKey Key { get; internal set; }

        /// <summary>
        /// Gets the value contained in the node.
        /// </summary>
        public TValue Value { get; internal set; }

        /// <summary>
        /// Gets or sets the parent node in the <see cref="OrderedMap{TKey, TValue}"/>.
        /// </summary>
        internal Node? Parent
        {
            get => _parent;
            set
            {
                _parent = value;
                if (Value != null)
                {
                    Value.Parent = value?.Value;
                }
            }
        }

        /// <summary>
        /// Gets or sets the left node in the <see cref="OrderedMap{TKey, TValue}"/>.
        /// </summary>
        internal Node? Left
        {
            get => _left;
            set
            {
                _left = value;
                if (Value != null)
                {
                    Value.Left = value?.Value;
                }
            }
        }

        /// <summary>
        /// Gets or sets the right node in the <see cref="OrderedMap{TKey, TValue}"/>.
        /// </summary>
        internal Node? Right
        {
            get => _right;
            set
            {
                _right = value;
                if (Value != null)
                {
                    Value.Right = value?.Value;
                }
            }
        }

        /// <summary>
        /// Gets or sets the color of the node.
        /// </summary>
        private NodeColor _color;

        private Node _left;
        private Node _right;
        private Node _parent;

        internal NodeColor Color
        {
            get => _color;
            set
            {
                _color = value;
                if (Value != null)
                {
                    Value.Color = (Color)value;
                }
            }
        }

        /// <summary>
        /// Gets the previous node in the <see cref="OrderedMap{TKey, TValue}"/>.
        /// <br/>O(log n) operation.
        /// </summary>
        public Node? Previous
        {
            get
            {
                Node? node;
                if (Left == null)
                {
                    node = this;
                    Node? p = Parent;
                    while (p != null && node == p.Left)
                    {
                        node = p;
                        p = p.Parent;
                    }

                    return p;
                }

                node = Left;
                while (node.Right != null)
                {
                    node = node.Right;
                }

                return node;
            }
        }

        /// <summary>
        /// Gets the next node in the <see cref="OrderedMap{TKey, TValue}"/>
        /// <br/>O(log n) operation.
        /// </summary>
        public Node? Next
        {
            get
            {
                Node? node;
                if (Right == null)
                {
                    node = this;
                    Node? p = Parent;
                    while (p != null && node == p.Right)
                    {
                        node = p;
                        p = p.Parent;
                    }

                    return p;
                }

                node = Right;
                while (node.Left != null)
                {
                    node = node.Left;
                }

                return node;
            }
        }

        internal static bool IsNonNullBlack(Node? node) => node != null && node.IsBlack;

        internal static bool IsNonNullRed(Node? node) => node != null && node.IsRed;

        internal static bool IsNullOrBlack(Node? node) => node == null || node.IsBlack;

        internal bool IsBlack => Color == NodeColor.Black;

        internal bool IsRed => Color == NodeColor.Red;

        internal bool IsUnused => Color == NodeColor.Unused;

        internal bool IsLinkedList => Color == NodeColor.LinkedList;

        public override string ToString() => Color + ": " + Value;

        internal void ColorBlack() => Color = NodeColor.Black;

        internal void ColorRed() => Color = NodeColor.Red;

        internal void Clear()
        {
            Key = default(TKey)!;
            Value = default(TValue)!;
            Parent = null;
            Left = null;
            Right = null;
            Color = NodeColor.Unused;
        }

        internal void Reset(TKey key, TValue value, NodeColor color)
        {
            Key = key;
            Value = value;
            Parent = null;
            Left = null;
            Right = null;
            Color = color;
        }
    }

    #endregion

    public Node? Root;
    private int _version;
    private KeyCollection? _keys;
    private ValueCollection? _values;

    /// <summary>
    /// Gets the number of nodes actually contained in the <see cref="OrderedMap{TKey, TValue}"/>.
    /// </summary>
    public int Count { get; private set; }

    public int CompareFactor { get; }

    public IComparer<TKey> Comparer { get; private set; }


    /// <summary>
    /// Initializes a new instance of the <see cref="OrderedMap{TKey, TValue}"/> class.
    /// </summary>
    /// <param name="reverse">true to reverses the comparison provided by the comparer. </param>
    public OrderedMap(bool reverse = false)
    {
        CompareFactor = reverse ? -1 : 1;
        Comparer = Comparer<TKey>.Default;
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="OrderedMap{TKey, TValue}"/> class.
    /// </summary>
    /// <param name="comparer">The default comparer to use for comparing objects.</param>
    /// <param name="reverse">true to reverses the comparison provided by the comparer. </param>
    public OrderedMap(IComparer<TKey> comparer, bool reverse = false)
    {
        CompareFactor = reverse ? -1 : 1;
        Comparer = comparer ?? Comparer<TKey>.Default;
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="OrderedMap{TKey, TValue}"/> class.
    /// </summary>
    /// <param name="dictionary">The IDictionary implementation to copy to a new collection.</param>
    /// <param name="reverse">true to reverses the comparison provided by the comparer. </param>
    public OrderedMap(IDictionary<TKey, TValue> dictionary, bool reverse = false)
        : this(dictionary, Comparer<TKey>.Default, reverse)
    {
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="OrderedMap{TKey, TValue}"/> class.
    /// </summary>
    /// <param name="dictionary">The IDictionary implementation to copy to a new collection.</param>
    /// <param name="comparer">The default comparer to use for comparing objects.</param>
    /// <param name="reverse">true to reverses the comparison provided by the comparer. </param>
    public OrderedMap(IDictionary<TKey, TValue> dictionary, IComparer<TKey> comparer, bool reverse = false)
    {
        CompareFactor = reverse ? -1 : 1;
        Comparer = comparer ?? Comparer<TKey>.Default;

        foreach (var x in dictionary)
        {
            Add(x.Key, x.Value);
        }
    }

    /// <summary>
    /// Gets the first node in the <see cref="OrderedMap{TKey, TValue}"/>.
    /// </summary>
    public Node? First
    {
        get
        {
            if (Root == null)
            {
                return null;
            }

            var node = Root;
            while (node.Left != null)
            {
                node = node.Left;
            }

            return node;
        }
    }

    /// <summary>
    /// Gets the last node in the <see cref="OrderedMap{TKey, TValue}"/>. O(log n) operation.
    /// </summary>
    public Node? Last
    {
        get
        {
            if (Root == null)
            {
                return null;
            }

            var node = Root;
            while (node.Right != null)
            {
                node = node.Right;
            }

            return node;
        }
    }

    #region Enumerator

    public Enumerator GetEnumerator() => new Enumerator(this, Enumerator.KeyValuePair);

    IEnumerator<KeyValuePair<TKey, TValue>> IEnumerable<KeyValuePair<TKey, TValue>>.GetEnumerator() =>
        new Enumerator(this, Enumerator.KeyValuePair);

    IEnumerator IEnumerable.GetEnumerator() => new Enumerator(this, Enumerator.KeyValuePair);

    public struct Enumerator : IEnumerator<KeyValuePair<TKey, TValue>>, IDictionaryEnumerator
    {
        internal const int KeyValuePair = 1;
        internal const int DictEntry = 2;

        private readonly OrderedMap<TKey, TValue> _map;
        private readonly int _version;
        private readonly int _getEnumeratorRetType;
        private Node? _node;
        private TKey? _key;
        private TValue? _value;

        internal Enumerator(OrderedMap<TKey, TValue> map, int getEnumeratorRetType)
        {
            _map = map;
            _version = _map._version;
            _getEnumeratorRetType = getEnumeratorRetType;
            _node = _map.First;
            _key = default;
            _value = default;
        }

        public void Dispose()
        {
            _node = null;
            _key = default;
            _value = default;
        }

        public bool MoveNext()
        {
            if (_version != _map._version)
            {
                throw ThrowVersionMismatch();
            }

            if (_node == null)
            {
                _key = default(TKey)!;
                _value = default(TValue)!;
                return false;
            }

            _key = _node.Key;
            _value = _node.Value;
            _node = _node.Next;
            return true;
        }

        DictionaryEntry IDictionaryEnumerator.Entry => new DictionaryEntry(_key!, _value!);

        object IDictionaryEnumerator.Key => _key!;

        object IDictionaryEnumerator.Value => _value!;

        public KeyValuePair<TKey, TValue> Current => new KeyValuePair<TKey, TValue>(_key!, _value!);

        object? IEnumerator.Current
        {
            get
            {
                if (_getEnumeratorRetType == DictEntry)
                {
                    return new DictionaryEntry(_key!, _value!);
                }

                return new KeyValuePair<TKey, TValue>(_key!, _value!);
            }
        }

        void IEnumerator.Reset() => Reset();

        internal void Reset()
        {
            if (_version != _map._version)
            {
                throw ThrowVersionMismatch();
            }

            _node = _map.First;
            _key = default;
            _value = default;
        }

        private static Exception ThrowVersionMismatch()
        {
            throw new InvalidOperationException("Collection was modified after the enumerator was instantiated.'");
        }
    }

    #endregion

    #region ICollection

    bool ICollection.IsSynchronized => false;

    object ICollection.SyncRoot => this;

    void ICollection.CopyTo(Array array, int index)
    {
        if (array == null)
        {
            throw new ArgumentNullException(nameof(array));
        }

        if (array.Rank != 1)
        {
            throw new ArgumentException(nameof(array));
        }

        if (array.GetLowerBound(0) != 0)
        {
            throw new ArgumentException(nameof(array));
        }

        if (index < 0 || index > array.Length)
        {
            throw new ArgumentOutOfRangeException(nameof(index));
        }

        if (array.Length - index < Count)
        {
            throw new ArgumentException();
        }

        var node = First;
        if (array is KeyValuePair<TKey, TValue>[] keyValuePairArray)
        {
            for (var i = 0; i < Count; i++)
            {
                keyValuePairArray[i + index] = new KeyValuePair<TKey, TValue>(node!.Key, node!.Value);
                node = node.Next;
            }
        }
        else
        {
            object[]? objects = array as object[];
            if (objects == null)
            {
                throw new ArgumentException(nameof(array));
            }

            try
            {
                for (int i = 0; i < Count; i++)
                {
                    objects[i + index] = new KeyValuePair<TKey, TValue>(node!.Key, node!.Value);
                    node = node.Next;
                }
            }
            catch (ArrayTypeMismatchException)
            {
                throw new ArgumentException(nameof(array));
            }
        }
    }

    #endregion

    #region IDictionary

    object? IDictionary.this[object? key]
    {
        get
        {
            return key switch
            {
                null when TryGetValue(default!, out var value) => value,
                TKey k when TryGetValue(k, out var value) => value,
                _ => null
            };
        }

        set => this[((TKey)key)!] = (TValue)value!;
    }

    bool IDictionary.IsFixedSize => false;

    bool IDictionary.IsReadOnly => false;

    ICollection IDictionary.Keys => Keys;

    ICollection IDictionary.Values => Values;

    void IDictionary.Add(object key, object? value) => Add((TKey)key, (TValue)value!);

    bool IDictionary.Contains(object key)
    {
        return key switch
        {
            null => ContainsKey(default),
            TKey k => ContainsKey(k),
            _ => false
        };
    }

    IDictionaryEnumerator IDictionary.GetEnumerator() => new Enumerator(this, Enumerator.DictEntry);

    void IDictionary.Remove(object key)
    {
        switch (key)
        {
            case null:
                Remove(default);
                break;
            case TKey k:
                Remove(k);
                break;
        }
    }

    #endregion

    #region IDictionary<TKey, TValue>

    IEnumerable<TKey> IReadOnlyDictionary<TKey, TValue>.Keys => Keys;

    IEnumerable<TValue> IReadOnlyDictionary<TKey, TValue>.Values => Values;

    ICollection<TKey> IDictionary<TKey, TValue>.Keys => Keys;

    ICollection<TValue> IDictionary<TKey, TValue>.Values => Values;

    void IDictionary<TKey, TValue>.Add(TKey key, TValue value) => Add(key, value);

    #endregion

    #region ICollection<KeyValuePair<TKey,TValue>>

    bool ICollection<KeyValuePair<TKey, TValue>>.IsReadOnly => false;

    void ICollection<KeyValuePair<TKey, TValue>>.Add(KeyValuePair<TKey, TValue> item) => Add(item.Key, item.Value);

    bool ICollection<KeyValuePair<TKey, TValue>>.Contains(KeyValuePair<TKey, TValue> item)
    {
        var node = FindNode(item.Key);
        return node != null && EqualityComparer<TValue>.Default.Equals(node.Value, item.Value);
    }

    void ICollection<KeyValuePair<TKey, TValue>>.CopyTo(KeyValuePair<TKey, TValue>[] array, int index) =>
        ((ICollection)this).CopyTo(array, index);

    bool ICollection<KeyValuePair<TKey, TValue>>.Remove(KeyValuePair<TKey, TValue> item)
    {
        var node = FindNode(item.Key);
        if (node == null || !EqualityComparer<TValue>.Default.Equals(node.Value, item.Value))
        {
            return false;
        }

        RemoveNode(node);
        return true;
    }

    #endregion

    #region KeyValueCollection

    public KeyCollection Keys => _keys ??= new KeyCollection(this);

    public ValueCollection Values => _values ??= new ValueCollection(this);

    public sealed class KeyCollection : ICollection<TKey>, ICollection, IReadOnlyCollection<TKey>
    {
        private readonly OrderedMap<TKey, TValue> _map;

        public KeyCollection(OrderedMap<TKey, TValue> map)
        {
            _map = map ?? throw new ArgumentNullException(nameof(map));
        }

        public Enumerator GetEnumerator() => new Enumerator(_map);

        IEnumerator<TKey> IEnumerable<TKey>.GetEnumerator() => new Enumerator(_map);

        IEnumerator IEnumerable.GetEnumerator() => new Enumerator(_map);

        public void CopyTo(TKey[] array, int index)
        {
            if (array == null)
            {
                throw new ArgumentNullException(nameof(array));
            }

            if (index < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }

            if (array.Length - index < Count)
            {
                throw new ArgumentException();
            }

            var node = _map.First;
            while (node != null)
            {
                array[index++] = node.Key;
                node = node.Next;
            }
        }

        void ICollection.CopyTo(Array array, int index)
        {
            if (array == null)
            {
                throw new ArgumentNullException(nameof(array));
            }

            if (array.Rank != 1)
            {
                throw new ArgumentException(nameof(array));
            }

            if (array.GetLowerBound(0) != 0)
            {
                throw new ArgumentException(nameof(array));
            }

            if (index < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }

            if (array.Length - index < _map.Count)
            {
                throw new ArgumentException();
            }

            if (array is TKey[] keys)
            {
                CopyTo(keys, index);
            }
            else
            {
                try
                {
                    var objects = (object[])array;
                    var node = _map.First;
                    while (node != null)
                    {
                        objects[index++] = node.Key!;
                        node = node.Next;
                    }
                }
                catch (ArrayTypeMismatchException)
                {
                    throw new ArgumentException(nameof(array));
                }
            }
        }

        public int Count => _map.Count;

        bool ICollection<TKey>.IsReadOnly => true;

        void ICollection<TKey>.Add(TKey item) => throw new NotSupportedException();

        void ICollection<TKey>.Clear() => throw new NotSupportedException();

        bool ICollection<TKey>.Contains(TKey item) => _map.ContainsKey(item);

        bool ICollection<TKey>.Remove(TKey item) => throw new NotSupportedException();

        bool ICollection.IsSynchronized => false;

        object ICollection.SyncRoot => ((ICollection)_map).SyncRoot;

        public struct Enumerator : IEnumerator<TKey>, IEnumerator
        {
            private IEnumerator<KeyValuePair<TKey, TValue>> _mapEnum;

            internal Enumerator(OrderedMap<TKey, TValue> map)
            {
                _mapEnum = map.GetEnumerator();
            }

            public void Dispose() => _mapEnum.Dispose();

            public bool MoveNext() => _mapEnum.MoveNext();

            public TKey Current => _mapEnum.Current.Key;

            object? IEnumerator.Current => Current;

            void IEnumerator.Reset() => _mapEnum.Reset();
        }
    }

    public sealed class ValueCollection : ICollection<TValue>, ICollection, IReadOnlyCollection<TValue>
    {
        private readonly OrderedMap<TKey, TValue> _map;

        public ValueCollection(OrderedMap<TKey, TValue> map)
        {
            _map = map ?? throw new ArgumentNullException(nameof(map));
        }

        public Enumerator GetEnumerator() => new Enumerator(_map);

        IEnumerator<TValue> IEnumerable<TValue>.GetEnumerator() => new Enumerator(_map);

        IEnumerator IEnumerable.GetEnumerator() => new Enumerator(_map);

        public void CopyTo(TValue[] array, int index)
        {
            if (array == null)
            {
                throw new ArgumentNullException(nameof(array));
            }

            if (index < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }

            if (array.Length - index < Count)
            {
                throw new ArgumentException();
            }

            var node = _map.First;
            while (node != null)
            {
                array[index++] = node.Value;
                node = node.Next;
            }
        }

        void ICollection.CopyTo(Array array, int index)
        {
            if (array == null)
            {
                throw new ArgumentNullException(nameof(array));
            }

            if (array.Rank != 1)
            {
                throw new ArgumentException(nameof(array));
            }

            if (array.GetLowerBound(0) != 0)
            {
                throw new ArgumentException(nameof(array));
            }

            if (index < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }

            if (array.Length - index < _map.Count)
            {
                throw new ArgumentException();
            }

            if (array is TValue[] values)
            {
                CopyTo(values, index);
            }
            else
            {
                try
                {
                    var objects = (object?[])array;
                    var node = _map.First;
                    while (node != null)
                    {
                        objects[index++] = node.Value;
                        node = node.Next;
                    }
                }
                catch (ArrayTypeMismatchException)
                {
                    throw new ArgumentException(nameof(array));
                }
            }
        }

        public int Count => _map.Count;

        bool ICollection<TValue>.IsReadOnly => true;

        void ICollection<TValue>.Add(TValue item) => throw new NotSupportedException();

        void ICollection<TValue>.Clear() => throw new NotSupportedException();

        bool ICollection<TValue>.Contains(TValue item)
        {
            return _map.ContainsValue(item);
        }

        bool ICollection<TValue>.Remove(TValue item) => throw new NotSupportedException();

        bool ICollection.IsSynchronized => false;

        object ICollection.SyncRoot => ((ICollection)_map).SyncRoot;

        public struct Enumerator : IEnumerator<TValue>
        {
            private IEnumerator<KeyValuePair<TKey, TValue>> _mapEnum;

            internal Enumerator(OrderedMap<TKey, TValue> map)
            {
                _mapEnum = map.GetEnumerator();
            }

            public void Dispose() => _mapEnum.Dispose();

            public bool MoveNext() => _mapEnum.MoveNext();

            public TValue Current => _mapEnum.Current.Value;

            object? IEnumerator.Current => Current;

            void IEnumerator.Reset() => _mapEnum.Reset();
        }
    }

    #endregion

    #region Main

    public TValue this[TKey key]
    {
        get
        {
            var node = FindNode(key);
            if (node == null)
            {
                throw new KeyNotFoundException();
            }

            return node.Value;
        }

        set
        {
            var node = FindNode(key);
            if (node == null)
            {
                Add(key, value);
            }
            else
            {
                node.Value = value;
            }
        }
    }

    public bool ContainsKey(TKey? key) => FindNode(key) != null;

    public bool ContainsValue(TValue value)
    {
        var found = false;

        if (value == null)
        {
            var node = First;
            while (node != null)
            {
                if (node.Value == null)
                {
                    found = true;
                    break;
                }

                node = node.Next;
            }
        }
        else
        {
            var comparer = EqualityComparer<TValue>.Default;
            var node = First;
            while (node != null)
            {
                if (comparer.Equals(node.Value, value))
                {
                    found = true;
                    break;
                }

                node = node.Next;
            }
        }

        return found;
    }

#pragma warning disable CS8767 // Nullability of reference types in type of parameter doesn't match implicitly implemented member (possibly because of nullability attributes).
    public bool TryGetValue(TKey? key, [MaybeNullWhen(false)] out TValue value)
#pragma warning restore CS8767 // Nullability of reference types in type of parameter doesn't match implicitly implemented member (possibly because of nullability attributes).
    {
        if (key == null)
        {
            throw new ArgumentNullException(nameof(key));
        }

        var node = FindNode(key);
        if (node == null)
        {
            value = default;
            return false;
        }

        value = node.Value;
        return true;
    }

    /// <summary>
    /// Removes all elements from a collection.
    /// </summary>
    public void Clear()
    {
        Root = null;
        _version = 0;
        Count = 0;
    }

    /// <summary>
    /// Copies the elements of the collection to the specified array of KeyValuePair structures, starting at the specified index.
    /// </summary>
    /// <param name="array">The one-dimensional array of KeyValuePair structures that is the destination of the elements.</param>
    /// <param name="index">The zero-based index in array at which copying begins.</param>
    public void CopyTo(KeyValuePair<TKey, TValue>[] array, int index) => ((ICollection)this).CopyTo(array, index);

    /// <summary>
    /// Removes a specified item from a collection.
    /// <br/>O(log n) operation.
    /// </summary>
    /// <param name="key">The element to remove.</param>
    /// <returns>true if the element is found and successfully removed.</returns>
    public bool Remove(TKey? key)
    {
        var p = FindNode(key);
        if (p == null)
        {
            return false;
        }

        RemoveNode(p);
        return true;
    }

    /// <summary>
    /// Adds an element to a collection. If the element is already in the set, this method returns the stored element without creating a new node, and sets newlyAdded to false.
    /// <br/>O(log n) operation.
    /// </summary>
    /// <param name="key">The key of the element to add.</param>
    /// <param name="value">The value of the element to add.</param>
    /// <returns>node: the added <see cref="OrderedMap{TKey, TValue}.Node"/>.<br/>
    /// newlyAdded: true if the node is created.</returns>
    public (Node node, bool newlyAdded) Add(TKey? key, TValue value) => Probe(key, value, null);

    /// <summary>
    /// Adds an element to a collection. If the element is already in the set, this method returns the stored element without creating a new node, and sets newlyAdded to false.
    /// <br/>O(log n) operation.
    /// </summary>
    /// <param name="key">The key of the element to add.</param>
    /// <param name="value">The value of the element to add.</param>
    /// <param name="reuse">Reuse a node to avoid memory allocation.</param>
    /// <returns>node: the added <see cref="OrderedMap{TKey, TValue}.Node"/>.<br/>
    /// newlyAdded: true if the node is created.</returns>
    public (Node node, bool newlyAdded) Add(TKey key, TValue value, Node reuse) => Probe(key, value, reuse);

    /// <summary>
    /// Updates the node's key with the specified key. Removes the node and inserts in the correct position if necessary.
    /// <br/>O(log n) operation.
    /// </summary>
    /// <param name="node">The <see cref="OrderedMap{TKey, TValue}.Node"/> to change the key.</param>
    /// <param name="key">The key to set.</param>
    /// <returns>true if the key is changed.</returns>
    public bool SetNodeKey(Node node, TKey key)
    {
        if (Comparer.Compare(node.Key, key) == 0)
        {
            // Identical
            return false;
        }

        var value = node.Value;
        RemoveNode(node);
        Probe(key, value, node);
        return true;
    }

    /// <summary>
    /// Updates the node's value with the specified value.
    /// <br/>O(1) operation.
    /// </summary>
    /// <param name="node">The <see cref="OrderedMap{TKey, TValue}.Node"/> to change the value.</param>
    /// <param name="value">The value to set.</param>
    public void SetNodeValue(Node node, TValue value) => node.Value = value;

    /// <summary>
    /// Removes a specified node from the collection"/>.
    /// <br/>O(log n) operation.
    /// </summary>
    /// <param name="node">The <see cref="OrderedMap{TKey, TValue}.Node"/> to remove.</param>
    public void RemoveNode(Node node)
    {
        Node? f; // Node to fix.
        var dir = 0;

        var originalColor = node.Color;
        if (node.Color == NodeColor.Unused)
        {
            // empty
            return;
        }

        f = node.Parent;
        if (node.Parent == null)
        {
            dir = 0;
        }
        else if (node.Parent.Left == node)
        {
            dir = -1;
        }
        else if (node.Parent.Right == node)
        {
            dir = 1;
        }

        _version++;
        Count--;

        if (node.Left == null)
        {
            TransplantNode(node.Right, node);
        }
        else if (node.Right == null)
        {
            TransplantNode(node.Left, node);
        }
        else
        {
            // Minimum
            Node? m = node.Right;
            while (m.Left != null)
            {
                m = m.Left;
            }

            originalColor = m.Color;
            if (m.Parent == node)
            {
                f = m;
                dir = 1;
            }
            else
            {
                f = m.Parent;
                dir = -1;

                TransplantNode(m.Right, m);
                m.Right = node.Right;
                m.Right.Parent = m;
            }

            TransplantNode(m, node);
            m.Left = node.Left;
            m.Left.Parent = m;
            m.Color = node.Color;
        }

        if (originalColor == NodeColor.Red || f == null)
        {
            node.Clear();
            Root?.ColorBlack();
            return;
        }

        while (true)
        {
            Node? s;
            if (dir < 0)
            {
                s = f.Right;
                if (Node.IsNonNullRed(s))
                {
                    s!.ColorBlack();
                    f.ColorRed();
                    RotateLeft(f);
                    s = f.Right;
                }

                // s is null or black
                if (s == null)
                {
                    // loop
                }
                else if (Node.IsNullOrBlack(s.Left) && Node.IsNullOrBlack(s.Right))
                {
                    s.ColorRed();
                    // loop
                }
                else
                {
                    // s is black and one of children is red.
                    if (Node.IsNonNullRed(s.Left))
                    {
                        s.Left!.ColorBlack();
                        s.ColorRed();
                        RotateRight(s);
                        s = f.Right;
                    }

                    s!.Color = f.Color;
                    f.ColorBlack();
                    s.Right!.ColorBlack();
                    RotateLeft(f);
                    break;
                }
            }
            else
            {
                s = f.Left;
                if (Node.IsNonNullRed(s))
                {
                    s!.ColorBlack();
                    f.ColorRed();
                    RotateRight(f);
                    s = f.Left;
                }

                // s is null or black
                if (s == null)
                {
                    // loop
                }
                else if (Node.IsNullOrBlack(s.Left) && Node.IsNullOrBlack(s.Right))
                {
                    s.ColorRed();
                    // loop
                }
                else
                {
                    // s is black and one of children is red.
                    if (Node.IsNonNullRed(s.Right))
                    {
                        s.Right!.ColorBlack();
                        s.ColorRed();
                        RotateLeft(s);
                        s = f.Left;
                    }

                    s!.Color = f.Color;
                    f.ColorBlack();
                    s.Left!.ColorBlack();
                    RotateRight(f);
                    break;
                }
            }

            if (f.IsRed || f.Parent == null)
            {
                f.ColorBlack();
                break;
            }

            if (f == f.Parent.Left)
            {
                dir = -1;
            }
            else
            {
                dir = 1;
            }

            f = f.Parent;
        }

        node.Clear();
    }

    /// <summary>
    /// Searches a tree for the specific value.
    /// </summary>
    /// <param name="target">The node to search.</param>
    /// <param name="key">The value to search for.</param>
    /// <returns>cmp: -1 => left, 0 and leaf is not null => found, 1 => right.
    /// leaf: the node with the specific value if found, or the nearest parent node if not found.</returns>
    private (int cmp, Node? leaf) SearchNode(Node? target, TKey? key)
    {
        Node? x = target;
        Node? p = null;
        var cmp = 0;

        if (CompareFactor > 0)
        {
            if (key == null)
            {
                // key is null
                while (x != null)
                {
                    if (x.Key == null)
                    {
                        // null == null
                        return (0, x);
                    }

                    // null < not null
                    p = x;
                    cmp = -1;
                    x = x.Left;
                }
            }
            else if (Equals(Comparer, Comparer<TKey>.Default) && key is IComparable<TKey> ic)
            {
                // IComparable<TKey>
                while (x != null)
                {
                    cmp = ic.CompareTo(x.Key); // -1: 1st < 2nd, 0: equals, 1: 1st > 2nd
                    p = x;
                    switch (cmp)
                    {
                        case < 0:
                            x = x.Left;
                            break;
                        case > 0:
                            x = x.Right;
                            break;
                        default:
                            // Found
                            return (0, x);
                    }
                }
            }
            else
            {
                // IComparer<TKey>
                while (x != null)
                {
                    cmp = Comparer.Compare(key, x.Key); // -1: 1st < 2nd, 0: equals, 1: 1st > 2nd
                    p = x;
                    switch (cmp)
                    {
                        case < 0:
                            x = x.Left;
                            break;
                        case > 0:
                            x = x.Right;
                            break;
                        default:
                            // Found
                            return (0, x);
                    }
                }
            }
        }
        else
        {
            if (key == null)
            {
                // key is null
                while (x != null)
                {
                    if (x.Key == null)
                    {
                        // null == null
                        return (0, x);
                    }

                    // null > not null
                    p = x;
                    cmp = 1;
                    x = x.Right;
                }
            }
            else if (Equals(Comparer, Comparer<TKey>.Default) && key is IComparable<TKey> ic)
            {
                // IComparable<TKey>
                while (x != null)
                {
                    cmp = ic.CompareTo(x.Key); // -1: 1st < 2nd, 0: equals, 1: 1st > 2nd
                    p = x;
                    switch (cmp)
                    {
                        case > 0:
                            cmp = -1;
                            x = x.Left;
                            break;
                        case < 0:
                            cmp = 1;
                            x = x.Right;
                            break;
                        default:
                            // Found
                            return (0, x);
                    }
                }
            }
            else
            {
                // IComparer<TKey>
                while (x != null)
                {
                    cmp = Comparer.Compare(key, x.Key); // -1: 1st < 2nd, 0: equals, 1: 1st > 2nd
                    p = x;
                    switch (cmp)
                    {
                        case > 0:
                            cmp = -1;
                            x = x.Left;
                            break;
                        case < 0:
                            cmp = 1;
                            x = x.Right;
                            break;
                        default:
                            // Found
                            return (0, x);
                    }
                }
            }
        }

        return (cmp, p);
    }
    
    /// <summary>
    /// Searches for a <see cref="OrderedMap{TKey, TValue}.Node"/> with the specified value.
    /// </summary>
    /// <param name="key">The value to search in a collection.</param>
    /// <returns>The node with the specified value.</returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public bool TryFindNode(TKey key, out TValue node)
    {
        var result = SearchNode(Root, key);
        if (result.cmp != 0 || result.leaf is null)
        {
            node = default;
            return false;
        }
        node = result.leaf.Value;
        return true;
    }

    /// <summary>
    /// Searches for a <see cref="OrderedMap{TKey, TValue}.Node"/> with the specified value.
    /// </summary>
    /// <param name="key">The value to search in a collection.</param>
    /// <returns>The node with the specified value.</returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public Node? FindNode(TKey? key)
    {
        var result = SearchNode(Root, key);
        return result.cmp == 0 ? result.leaf : null;
    }

    /// <summary>
    /// Searches for the first <see cref="OrderedMap{TKey, TValue}.Node"/> with the key equal to or greater than the specified key (null: all nodes are less than the specified key).
    /// </summary>
    /// <param name="key">The key to search for.</param>
    /// <returns>The first <see cref="OrderedMap{TKey, TValue}.Node"/> with the key equal to or greater than the specified key (null: all nodes are less than the specified key).</returns>
    public Node? GetLowerBound(TKey? key)
    {
        var (cmp, p) = SearchNode(Root, key);

        return cmp switch
        {
            0 => p,
            < 0 => p,
            _ => p?.Next
        };
    }

    /// <summary>
    /// Searches for the last <see cref="OrderedMap{TKey, TValue}.Node"/> with the key equal to or lower than the specified key (null: all nodes are greater than the specified key).
    /// </summary>
    /// <param name="key">The key to search for.</param>
    /// <returns>The last <see cref="OrderedMap{TKey, TValue}.Node"/> with the key equal to or lower than the specified key (null: all nodes are greater than the specified key).</returns>
    public Node? GetUpperBound(TKey? key)
    {
        var (cmp, p) = SearchNode(Root, key);

        return cmp switch
        {
            0 => p,
            < 0 => p?.Previous,
            _ => p
        };
    }

    /// <summary>
    /// Gets <see cref="Node"/> whose keys are in the range from the lower bound to the upper bound.
    /// </summary>
    /// <param name="lower">Lower bound key.</param>
    /// <param name="upper">Upper bound key.</param>
    /// <returns>The lower and upper <see cref="Node"/>.</returns>
    public (Node? Lower, Node? Upper) GetRange(TKey? lower, TKey? upper)
    {
        var lowerNode = GetLowerBound(lower);
        if (lowerNode == null)
        {
            return (null, null);
        }

        var upperNode = GetUpperBound(upper);
        if (upperNode == null)
        {
            return (null, null);
        }

        return Comparer.Compare(lowerNode.Key, upperNode.Key) > 0 ? (null, null) : (lowerNode, upperNode);
    }

    /// <summary>
    /// Adds an element to the set. If the element is already in the set, this method returns the stored node without creating a new node.
    /// <br/>O(log n) operation.
    /// </summary>
    /// <param name="key">The element to add to the set.</param>
    /// <returns>node: the added <see cref="OrderedMap{TKey, TValue}.Node"/>.<br/>
    /// newlyAdded: true if the node is created.</returns>
    private (Node node, bool newlyAdded) Probe(TKey key, TValue value, Node? reuse)
    {
        Node? x = Root; // Traverses tree looking for insertion point.
        Node? p = null; // Parent of x; node at which we are rebalancing.
        var cmp = 0;

        (cmp, p) = SearchNode(Root, key);
        if (cmp == 0 && p != null)
        {
            // Found
            return (p, false);
        }

        _version++;
        Count++;

        Node n;
        if (reuse is { IsUnused: true })
        {
            reuse.Reset(key, value, NodeColor.Red);
            n = reuse;
        }
        else
        {
            n = new Node(key, value,
                NodeColor.Red); // Newly inserted node. // this.CreateNode(key, value, NodeColor.Red);
        }

        n.Parent = p;
        if (p != null)
        {
            if (cmp < 0)
            {
                p.Left = n;
            }
            else
            {
                p.Right = n;
            }
        }
        else
        {
            // Root
            Root = n;
            n.ColorBlack();
            return (n, true);
        }

        p = n;

#nullable disable
        while (p.Parent is { IsRed: true })
        {
            // p.Parent is not root (root is black), so p.Parent.Parent != null
            if (p.Parent == p.Parent.Parent.Right)
            {
                x = p.Parent.Parent.Left; // uncle
                if (x is { IsRed: true })
                {
                    x.ColorBlack();
                    p.Parent.ColorBlack();
                    p.Parent.Parent.ColorRed();
                    p = p.Parent.Parent; // loop
                }
                else
                {
                    if (p == p.Parent.Left)
                    {
                        p = p.Parent;
                        RotateRight(p);
                    }

                    p.Parent.ColorBlack();
                    p.Parent.Parent.ColorRed();
                    RotateLeft(p.Parent.Parent);
                    break;
                }
            }
            else
            {
                x = p.Parent.Parent.Right; // uncle

                if (x is { IsRed: true })
                {
                    x.ColorBlack();
                    p.Parent.ColorBlack();
                    p.Parent.Parent.ColorRed();
                    p = p.Parent.Parent; // loop
                }
                else
                {
                    if (p == p.Parent.Right)
                    {
                        p = p.Parent;
                        RotateLeft(p);
                    }

                    p.Parent.ColorBlack();
                    p.Parent.Parent.ColorRed();
                    RotateRight(p.Parent.Parent);
                    break;
                }
            }
        }
#nullable enable

        Root!.ColorBlack();
        return (n, true);
    }

    #endregion

    #region Validation

    /// <summary>
    /// Validate Red-Black Tree.
    /// </summary>
    /// <returns>true if the tree is valid.</returns>
    public bool Validate()
    {
        var result = true;
        result &= ValidateBst(Root);
        result &= ValidateBlackHeight(Root) >= 0;
        result &= ValidateColor(Root) == NodeColor.Black;

        return result;
    }

    private NodeColor ValidateColor(Node? node)
    {
        if (node == null)
        {
            return NodeColor.Black;
        }

        var color = node.Color;
        var leftColor = ValidateColor(node.Left);
        var rightColor = ValidateColor(node.Right);
        if (leftColor == NodeColor.Unused || rightColor == NodeColor.Unused)
        {
            // Error
            return NodeColor.Unused;
        }

        return color switch
        {
            NodeColor.Black => color,
            NodeColor.Red when leftColor == NodeColor.Black && rightColor == NodeColor.Black => color,
            _ => NodeColor.Unused
        };
    }

    private int ValidateBlackHeight(Node? node)
    {
        if (node == null)
        {
            return 0;
        }

        var leftHeight = ValidateBlackHeight(node.Left);
        var rightHeight = ValidateBlackHeight(node.Right);
        if (leftHeight < 0 || rightHeight < 0 || leftHeight != rightHeight)
        {
            // Invalid
            return -1;
        }

        return leftHeight + (node.IsBlack ? 1 : 0);
    }

    private bool ValidateBst(Node? node)
    {
        // Binary Search Tree
        if (node == null)
        {
            return true;
        }

        var result = true;

        if (node.Parent == null)
        {
            result &= Root == node;
        }

        if (node.Left != null)
        {
            result &= node.Left.Parent == node;
        }

        if (node.Right != null)
        {
            result &= node.Right.Parent == node;
        }

        result &= IsSmaller(node.Left, node.Key) && IsLarger(node.Right, node.Key);
        result &= ValidateBst(node.Left) && ValidateBst(node.Right);
        return result;
    }

    private bool IsSmaller(Node? node, TKey key)
    {
        // Node value is smaller than TKey value.
        if (node == null)
        {
            return true;
        }

        var cmp = Comparer.Compare(node.Key, key); // -1: 1st < 2nd, 0: equals, 1: 1st > 2nd
        return cmp == -1 && IsSmaller(node.Left, key) && IsSmaller(node.Right, key);
    }

    private bool IsLarger(Node? node, TKey key)
    {
        // Node value is larger than TKey value.
        if (node == null)
        {
            return true;
        }

        var cmp = Comparer.Compare(node.Key, key); // -1: 1st < 2nd, 0: equals, 1: 1st > 2nd
        return cmp == 1 && IsLarger(node.Left, key) && IsLarger(node.Right, key);
    }

    #endregion

    #region LowLevel

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private void TransplantNode(Node? node, Node destination)
    {
        // Transplant Node node to Node destination
        if (destination.Parent == null)
        {
            Root = node;
        }
        else if (destination == destination.Parent.Left)
        {
            destination.Parent.Left = node;
        }
        else
        {
            destination.Parent.Right = node;
        }

        if (node != null)
        {
            node.Parent = destination.Parent;
        }
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private void RotateLeft(Node x)
    {
        // checked
        var y = x.Right!;
        x.Right = y.Left;
        if (y.Left != null)
        {
            y.Left.Parent = x;
        }

        var p = x.Parent; // Parent of x
        y.Parent = p;
        if (p == null)
        {
            Root = y;
        }
        else if (x == p.Left)
        {
            p.Left = y;
        }
        else
        {
            p.Right = y;
        }

        y.Left = x;
        x.Parent = y;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private void RotateRight(Node x)
    {
        // checked
        var y = x.Left!;
        x.Left = y.Right;
        if (y.Right != null)
        {
            y.Right.Parent = x;
        }

        var p = x.Parent; // Parent of x
        y.Parent = p;
        if (p == null)
        {
            Root = y;
        }
        else if (x == p.Right)
        {
            p.Right = y;
        }
        else
        {
            p.Left = y;
        }

        y.Right = x;
        x.Parent = y;
    }

    #endregion
}
