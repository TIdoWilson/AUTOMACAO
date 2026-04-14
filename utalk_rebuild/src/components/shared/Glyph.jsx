function icon(name) {
  const stroke = { fill: 'none', stroke: 'currentColor', strokeWidth: 1.8, strokeLinecap: 'round', strokeLinejoin: 'round' }

  switch (name) {
    case 'chat':
      return <path {...stroke} d="M4.5 5.5h11v8h-7l-3 3v-3h-1z" />
    case 'user':
      return (
        <>
          <circle {...stroke} cx="10" cy="7" r="3" />
          <path {...stroke} d="M4.5 15.5c1.5-2.4 3.4-3.5 5.5-3.5s4 .9 5.5 3.5" />
        </>
      )
    case 'bot':
      return (
        <>
          <rect {...stroke} x="4.5" y="5.5" width="11" height="9" rx="2.5" />
          <path {...stroke} d="M10 3.5v2M7.5 10h.01M12.5 10h.01M7 14.5l-1.5 2M13 14.5l1.5 2" />
        </>
      )
    case 'send':
      return <path {...stroke} d="M4.5 10l11-5-3 11-2.5-4.5zM10 10l5.5-5" />
    case 'chart':
      return <path {...stroke} d="M5 14.5V10M10 14.5V6M15 14.5V8M4.5 15.5h11" />
    case 'gear':
      return (
        <>
          <circle {...stroke} cx="10" cy="10" r="2.5" />
          <path {...stroke} d="M10 4.5v1.5M10 14v1.5M15.5 10H14M6 10H4.5M13.9 6.1l-1 1M7.1 12.9l-1 1M13.9 13.9l-1-1M7.1 7.1l-1-1" />
        </>
      )
    case 'search':
      return (
        <>
          <circle {...stroke} cx="8.5" cy="8.5" r="4" />
          <path {...stroke} d="M11.7 11.7l3.3 3.3" />
        </>
      )
    case 'plus':
      return <path {...stroke} d="M10 5v10M5 10h10" />
    case 'filter':
      return <path {...stroke} d="M4.5 5.5h11l-4 4.5v4l-3 1v-5z" />
    case 'bell':
      return <path {...stroke} d="M6.5 13.5V9.3A3.5 3.5 0 0 1 10 5.8a3.5 3.5 0 0 1 3.5 3.5v4.2l1.2 1.2H5.3zM8 14.7a2 2 0 0 0 4 0" />
    case 'spark':
      return <path {...stroke} d="M10 4l1.2 3.3L14.5 8.5l-3.3 1.2L10 13l-1.2-3.3L5.5 8.5l3.3-1.2z" />
    case 'money':
      return <path {...stroke} d="M11.5 6.2c-.4-.5-1-.7-1.8-.7-1.1 0-1.8.6-1.8 1.5 0 .8.5 1.1 1.8 1.4 1.4.3 2.4.8 2.4 2.1 0 1.2-1 2-2.5 2-1 0-1.9-.3-2.5-.9M10 4.5v11" />
    case 'tag':
      return <path {...stroke} d="M4.5 7.5V4.5h3l7 7-3 3-7-7zM6.8 6.8h.01" />
    case 'building':
      return <path {...stroke} d="M6 15.5v-9l4-2 4 2v9M4.5 15.5h11M8 8.5h.01M12 8.5h.01M8 11.5h.01M12 11.5h.01" />
    case 'calendar':
      return (
        <>
          <rect {...stroke} x="4.5" y="6" width="11" height="9.5" rx="2" />
          <path {...stroke} d="M7 4.5v3M13 4.5v3M4.5 8.5h11" />
        </>
      )
    case 'smile':
      return (
        <>
          <circle {...stroke} cx="10" cy="10" r="5.5" />
          <path {...stroke} d="M7.6 8.5h.01M12.4 8.5h.01M7.5 11.5c.8 1 1.6 1.5 2.5 1.5s1.7-.5 2.5-1.5" />
        </>
      )
    case 'paperclip':
      return <path {...stroke} d="M7 10.5l4.5-4.5a2 2 0 1 1 2.8 2.8l-5.8 5.8a3 3 0 1 1-4.2-4.2l5.1-5.1" />
    case 'menu':
      return <path {...stroke} d="M5 7h10M5 10h10M5 13h10" />
    case 'x':
      return <path {...stroke} d="M6 6l8 8M14 6l-8 8" />
    case 'clock':
      return (
        <>
          <circle {...stroke} cx="10" cy="10" r="5.5" />
          <path {...stroke} d="M10 7v3.4l2.3 1.3" />
        </>
      )
    case 'check':
      return <path {...stroke} d="M5.5 10.5l2.6 2.6 6-6.1" />
    case 'note':
      return <path {...stroke} d="M5 4.5h10v11H5zM7.5 8h5M7.5 11h5" />
    case 'image':
      return (
        <>
          <rect {...stroke} x="4.5" y="5.5" width="11" height="9" rx="2" />
          <path {...stroke} d="M7.5 11l1.8-1.8 2.1 2.1 1.6-1.6 2 2M8 8h.01" />
        </>
      )
    case 'arrowDown':
      return <path {...stroke} d="M6.5 8.5L10 12l3.5-3.5" />
    default:
      return <circle {...stroke} cx="10" cy="10" r="5" />
  }
}

export function Glyph({ name }) {
  return (
    <svg className="glyph" viewBox="0 0 20 20" aria-hidden="true">
      {icon(name)}
    </svg>
  )
}
