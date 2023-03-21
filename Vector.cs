using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelToJson
{
    struct Vector2<T>
    {
        public T x;
        public T y;

        public Vector2(T x, T y)
        {
            this.x = x;
            this.y = y;
        }
    }

    struct Vector3<T>
    {
        public T x;
        public T y;
        public T z;
        public Vector3(T x, T y, T z)
        {
            this.x = x;
            this.y = y;
            this.z = z;
        }
    }
}
