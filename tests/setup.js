/**
 * Jest Setup File
 * This runs before each test file and sets up Chrome API mocks
 */

// Mock Chrome Storage API
const mockStorage = {
  local: {
    _data: {},
    get: jest.fn((keys) => {
      return Promise.resolve(
        keys === null
          ? { ...mockStorage.local._data }
          : typeof keys === 'string'
            ? { [keys]: mockStorage.local._data[keys] }
            : keys.reduce((acc, key) => {
                acc[key] = mockStorage.local._data[key];
                return acc;
              }, {})
      );
    }),
    set: jest.fn((data) => {
      Object.assign(mockStorage.local._data, data);
      return Promise.resolve();
    }),
    remove: jest.fn((keys) => {
      const keyArray = Array.isArray(keys) ? keys : [keys];
      keyArray.forEach(key => delete mockStorage.local._data[key]);
      return Promise.resolve();
    }),
    clear: jest.fn(() => {
      mockStorage.local._data = {};
      return Promise.resolve();
    })
  },
  session: {
    _data: {},
    get: jest.fn((keys) => {
      return Promise.resolve(
        keys === null
          ? { ...mockStorage.session._data }
          : typeof keys === 'string'
            ? { [keys]: mockStorage.session._data[keys] }
            : keys.reduce((acc, key) => {
                acc[key] = mockStorage.session._data[key];
                return acc;
              }, {})
      );
    }),
    set: jest.fn((data) => {
      Object.assign(mockStorage.session._data, data);
      return Promise.resolve();
    })
  }
};

// Mock Chrome Runtime API
const mockRuntime = {
  sendMessage: jest.fn(() => Promise.resolve()),
  onMessage: {
    addListener: jest.fn(),
    removeListener: jest.fn()
  }
};

// Mock Chrome Tabs API
const mockTabs = {
  query: jest.fn(() => Promise.resolve([])),
  update: jest.fn(() => Promise.resolve())
};

// Mock Chrome Windows API
const mockWindows = {
  update: jest.fn(() => Promise.resolve())
};

// Assign to global chrome object
global.chrome = {
  storage: mockStorage,
  runtime: mockRuntime,
  tabs: mockTabs,
  windows: mockWindows
};

// Helper to reset all mocks between tests
global.resetChromeMocks = () => {
  mockStorage.local._data = {};
  mockStorage.session._data = {};
  jest.clearAllMocks();
};

// Helper to pre-populate storage for tests
global.setMockStorageData = (data) => {
  mockStorage.local._data = { ...data };
};

// Reset before each test
beforeEach(() => {
  global.resetChromeMocks();
});
